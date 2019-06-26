'Created by Matthew Hull on 4/27/04

'This script will get the information needed to setup SWS in LDAP mode

Option Explicit

On Error Resume Next

Dim objRootDSE, strMessage, strAdmin, strPort, strDomainDN, strIP
Dim objconnection, objCommand, objRecordSet, strADSPath, objGroup
Dim objUser, bolAdminFound

'Create a RootDSE Object, this will be used to connect to the domain
Set objRootDSE = GetObject("LDAP://rootDSE")

'Exit the script if a Domain isn't located
If Err Then
   MsgBox "You are not connected to a Domain.",vbCritical,"Domain Not Found"
   Err.Clear
   WScript.Quit
End If

'Set some of the constants
strIP = "Enter the IP of a Domain Controller"
strPort = "389"
strDomainDN = objRootDSE.Get("DefaultNamingContext")

'********************************************************************************
'This section will try to get the administrators DN

'Establish a connection to Active Directory using ActiveX Data Object
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADSDSOObject;"

'Create the command object and attach it to the connection object
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

'objCommand defines the search base (in this case the whole domain) 
'a filter (All object of type Group) and some attributes 
'associated with the returned objects (CN and DistinguishedName)
objCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
">;(&(objectClass=group));CN,DistinguishedName"

'Initiate the LDAP query and return results to a RecordSet object.
Set objRecordSet = objCommand.Execute

'This will be used to determine if the administrator account was found
bolAdminFound = False

'Loop Thru each member of the record set
While Not objRecordset.EOF
   If uCase(objRecordset.Fields("CN")) = "DOMAIN ADMINS" Then
      
      'Build the Group path and create the Group object
      strADSPath = "LDAP://" & objRecordset.Fields("DistinguishedName")
      Set objGroup = GetObject(strADSPath)
      
      'Loop thru each user in the group and find that administrator
      For Each objUser in objGroup.Members
         If UCase(objUser.SamAccountName) = "ADMINISTRATOR" Then
            strAdmin = objUser.DistinguishedName
            bolAdminFound = True
         ElseIf strAdmin = "" Then
            strAdmin = objUser.DistinguishedName
         End If         
      Next
   End If
   
   'Move to the next record
   objRecordset.MoveNext
Wend

'Check to see if the administrator account was found
If Not bolAdminFound Then
   MsgBox "The Administrator account was not found on the domain.  This " & _
   "Probably because the account has been renamed.  The first account in "& _
   "the Domain Admins group will be used instead.  Feel free to change it.", _
   vbOkOnly,"Administrator Account Not Found"
End If

'********************************************************************************

'Build the message to the user
strMessage = "Server Name/Address: " & strIP & vbCRLF
strMessage = strMessage & "Server Port Number: " & strPort & vbCRLF
strMessage = strMessage & "Administrator Name: " & strAdmin & vbCRLF
strMessage = strMessage & "Administrator Password: ************" & vbCRLF
strMessage = strMessage & "Root Node DN: " & strDomainDN & vbCRLF

'Display the message to the user
MsgBox strMessage,vbOkOnly,"SWS LDAP Information"

'Close Objects
Set objRootDSE = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRecordSet = Nothing