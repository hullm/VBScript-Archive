'Created By Matthew Hull on 5/12/04

'This script will remove the pause on a sync warning in Windows.
'This warning is caused by unsoppurted file types trying to sync.
'I.e. PST or MDB etc...

Option Explicit

On Error Resume Next

Dim WSHShell,strRegKey,strRegSetting

'Create the Shell object, this will be used to write to the registry
Set WSHShell = WScript.CreateObject("Wscript.Shell")

'1 - Pause on errors.
'2 - Pause on warnings.
'3 - Pause on errors and warnings.
'4 - Pause and display INFO.

strRegSetting = 1

'This is where the values will be stored in the registry
strRegKey="HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Syncmgr\"

'Write the values to the registry
WSHShell.RegWrite strRegKey & "KeepProgressLevel", strRegSetting, "REG_DWORD"

'Echo a message to the user when done
MsgBox "The synchronization pause has been removed",vbOkOnly,"Error Removed"