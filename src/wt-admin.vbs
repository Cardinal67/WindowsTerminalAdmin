' I had an issue with running the original script in some directories, so I had Claude 4.5 Sonnet rewrite this is account for that. 
' Thank you akopetsch for the original. Your work is greatly appreciated! 

' SPDX-License-Identifier: MIT
' Copyright (c) 2025 Alexander Kopetsch (Enhanced Version)

' Create objects
Set shell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

' Get the directory argument
If WScript.Arguments.Count > 0 Then
    directory = WScript.Arguments(0)
Else
    ' No argument provided, use user profile
    directory = wshShell.ExpandEnvironmentStrings("%USERPROFILE%")
End If

' Clean up the directory path
' Remove quotes if present
directory = Replace(directory, """", "")

' Handle root drive special case (D:\ becomes D:)
If Len(directory) = 3 And Right(directory, 2) = ":\" Then
    ' For root drives, keep the backslash for proper formatting
    ' but ensure no double backslashes
    directory = Left(directory, 3)
ElseIf Right(directory, 1) = "\" And Len(directory) > 3 Then
    ' Remove trailing backslash for non-root paths
    directory = Left(directory, Len(directory) - 1)
End If

' Verify the directory exists and is accessible
validDirectory = directory
If Not fso.FolderExists(directory) Then
    ' Try without the backslash for root drives
    If Len(directory) = 3 And Right(directory, 1) = "\" Then
        testDir = Left(directory, 2)
        If fso.FolderExists(testDir) Then
            validDirectory = testDir
        Else
            ' Fallback to user profile if directory doesn't exist
            validDirectory = wshShell.ExpandEnvironmentStrings("%USERPROFILE%")
        End If
    Else
        ' Fallback to user profile if directory doesn't exist
        validDirectory = wshShell.ExpandEnvironmentStrings("%USERPROFILE%")
    End If
End If

' Detect what's available and launch accordingly
On Error Resume Next

' First try: Windows Terminal (wt.exe)
wshShell.Run "where wt.exe", 0, True
If Err.Number = 0 Then
    ' Windows Terminal is available
    shell.ShellExecute "wt.exe", "-d """ & validDirectory & """", "", "runas", 1
Else
    ' Fallback: PowerShell directly
    Err.Clear
    ' Try PowerShell 7+ (pwsh.exe) first
    wshShell.Run "where pwsh.exe", 0, True
    If Err.Number = 0 Then
        shell.ShellExecute "pwsh.exe", "-NoExit -WorkingDirectory """ & validDirectory & """", "", "runas", 1
    Else
        ' Fallback to Windows PowerShell 5.1
        shell.ShellExecute "powershell.exe", "-NoExit -Command ""Set-Location '" & validDirectory & "'""", "", "runas", 1
    End If
End If

On Error GoTo 0
