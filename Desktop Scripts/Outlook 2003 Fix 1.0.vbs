'Created By Matthew Hull on some day a long time ago
'Documented on 4/25/04

'By default Outlook blocks extentions that are considered harmful.  This script will
'unblock the extentions that you list in strAllow

Dim WSHShell,strRegKey,strAllow

'*******************************************************************************************

'Setting the Value of the new string. Example "URL; EXE; COM"
strAllow="URL"

'*******************************************************************************************

'Create the Windows Scripting Host Shell Object, This allows for the Registry Write.
Set WSHShell = WScript.CreateObject("Wscript.Shell")

'Warns the users to exit Outlook
WSHShell.Popup "Please Close Outlook 2003" & Chr(10) & Chr(10) & _ 
   "You will need to restart your computer after this..."
 
'Setting the RegKey String to the String value I want to modify
strRegKey="HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Outlook\Security\Level1remove"

'Writting the RegKey with the new data.
WSHShell.RegWrite strRegKey, strAllow

'Let the user know when the process is complete
WSHShell.Popup "The string """ & strAllow & """" & " has been added to " & strRegKey & _
   Chr(10) & Chr(10) & "Please Restart Your computer..."

