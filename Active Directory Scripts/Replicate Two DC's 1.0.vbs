'Created by Matthew Hull on 7/21/04
'Last Updated 7/30/04

'This will replicate data from one domain controller to another.  It will prompt you for
'the name of each server.

Option Explicit

On Error Resume Next

Dim objIADSTools, objRootDSE, strSourceServer, strDestServer, strDomainDN, strSchemaDN
Dim strConfigDN, intResult, intError

'******************************************************************************************

'Create the required objects
Set objIADSTools = CreateObject("IADSTools.DCFunctions")

'Exit the script if it can't find the support tools
If Err Then
   MsgBox "Windows 2000/2003 Server Support Tools are not installed on this computer."  & _
   vbCRLF & "The script will now exit...",vbCritical,"Support Tools Missing"
   WScript.Quit
End If

'Create the RootDSE Object
Set objRootDSE = GetObject("LDAP://rootDSE")

'Exit if the user is not connected to a domain
If Err Then
   MsgBox "There was an error connecting to your domain.  Make sure you are " & _
   "connected to the network, or you might have to run this from a Domain " & _
   "Controller",vbOkOnly,"Cannot Connect to Domain"
   WScript.Quit
End If 

'******************************************************************************************

'Get the name of the source server
Do 
   strSourceServer = InputBox  ("Enter the source server name","Server Name"," ")
   If strSourceServer = "" Then
      WScript.Quit
   ElseIf strSourceServer = " " Then
      MsgBox "You must specify a server",vbOkOnly,"Server Required"
   End If
Loop Until strSourceServer <> " "

'Get the name of the destination server
Do 
   strDestServer = InputBox  ("Enter the destination server name","Server Name"," ")
   If strDestServer = "" Then
      WScript.Quit
   ElseIf strDestServer = " " Then
      MsgBox "You must specify a server",vbOkOnly,"Server Required"
   End If
Loop Until strDestServer <> " "

'******************************************************************************************

'Get the Default Naming Context
strDomainDN  = objRootDSE.Get("DefaultNamingContext")

'Get the Schema Naming Context
strSchemaDN = objRootDSE.Get("SchemaNamingContext")

'Get the Configuration Naming Context
strConfigDN = objRootDSE.Get("ConfigurationNamingContext")

'******************************************************************************************

'Replicate Default Naming Context
intResult = objIADSTools.ReplicaSync(CStr(strDestServer),CStr(strDomainDN), _
   CStr(strSourceServer))

'Check for errros
If intResult = -1 Then
   intError = 1
End If

'******************************************************************************************

'Replicate Schema Naming Context
intResult = objIADSTools.ReplicaSync(CStr(strDestServer),CStr(strSchemaDN), _
   CStr(strSourceServer))

'Check for errros
If intResult = -1 Then
   intError = 1
End If

'******************************************************************************************

'Replicate Configuration Naming Context
intResult = objIADSTools.ReplicaSync(CStr(strDestServer),CStr(strConfigDN), _
   CStr(strSourceServer))

'Check for errros
If intResult = -1 Then
   intError = 1
End If

'******************************************************************************************

If intError = 1 Then
   MsgBox "Error replicating from " & strSourceServer & " to " & strDestServer & "." & _
   vbCRLF & "Error: " & objIadsTools.LastErrorText,vbCritical,"Servers NOT Replicated"
Else
   MsgBox "Data replicated from " & strSourceServer & " to " & strDestServer & " on " & _
   Date & " at " & Time,vbOkOnly,"Replication Complete"
End If