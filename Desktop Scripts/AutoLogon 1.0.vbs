'Created By Matthew Hull on some day a long time ago
'Documented on 4/25/04

'This script will set a computer to auto logon

Dim WSHShell,strRegKey,strAllow

'Create the Shell object, this will be used to write to the registry
Set WSHShell = WScript.CreateObject("Wscript.Shell")

'This is used to turn on Auto Logon
strAutoAdminLogon="1"

'************************************************************************************
'This is the information that you need to change

strUserName= "username"
strPassword= "password"
strDomain ="Domain"

'************************************************************************************

'This is where the values will be stored in the registry
strRegKey="HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\"

'Write the values to the registry
WSHShell.RegWrite strRegKey & "AutoAdminLogon", strAutoAdminLogon
WSHShell.RegWrite strRegKey & "DefaultUserName", strUserName
WSHShell.RegWrite strRegKey & "DefaultPassword", strPassword
WSHShell.RegWrite strRegKey & "DefaultDomainName", strDomain

'Echo a message to the user when done
WScript.Echo "The user " & strUserName & " is now set to autologon."