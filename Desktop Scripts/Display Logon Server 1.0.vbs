'Created by Matthew Hull on 10/10/03
'Documented on 4/25/04

'This script will echo your logon server and username.

Option Explicit

Dim objWSHShell, objEnvironment, strVariable, strMessage, objRegExp
Dim strLogonServer, strUSerName

'Create the Shell Object, this will be used to get environmental variables
Set objWSHShell = CreateObject("WScript.Shell")

'Create a Regular Expression Object, this will be used to remove "\\"
'from the start of the logon server name
Set objRegExp = New RegExp

'Set the pattern for the regular expression
objRegExp.Pattern = "\\\\"

'Get the logon server name
strLogonServer = objWSHShell.ExpandEnvironmentStrings("%logonserver%")
strLogonServer = lcase(objRegExp.Replace(strLogonServer,""))

'Get the username
strUserName = lcase(objWSHShell.ExpandEnvironmentStrings("%username%"))

'Build the message that will display to the user
strMessage = "Your logon server is " & strLogonServer
strMessage = strMessage & " and you are logged in as " & strUserName & "."

'Check for a logon server, if one is missing change the message
If strLogonServer = "" Then
   strMessage = "You are running Windows 9x/ME, this program is designed "
   strMessage = strMessage & "to run in Windows NT/2k/XP"
End If

'Display the message to the user
MsgBox strMessage,vbOkOnly,"Logon Server"

'Close objects
Set objWSHShell = Nothing
Set objRegExp = Nothing