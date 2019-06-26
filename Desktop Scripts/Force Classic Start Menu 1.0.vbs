'Created by Matthew Hull on 8/9/04

'This sets an XP/2003 computer to use the classic interface for all users.

Option Explicit

On Error Resume Next

Dim objShell, strRegKey

'Create the Shell object, this will be used to write to the registry
Set objShell = WScript.CreateObject("Wscript.Shell")

'Set the root variables
strRegKey = "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\"

'Write data to the regiester
objShell.RegWrite strRegKey & "NoSimpleStartMenu", 1, "REG_DWORD"

'Display a message to the user when done.
MsgBox "The computer will now use the classic start menu for all users.",vbOkOnly,"Done"

'Close open objects
Set objShell = Nothing