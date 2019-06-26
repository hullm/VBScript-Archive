'Created by Matthew Hull on 1/28/05

'This script will try and count all your mailboxes.  It gathers the mail server
'address by looping through all the users in the Domain Admins group and looking
'at there msExchHomeServerName property.  If it isn't found the script will exit.
'If it is found it will connect to the mail server and count the mailboxes.

Option Explicit

On Error Resume Next

Dim objRootDSE, objConnection, objCommand, objRecordSet, strMailServerAddress
Dim strADSPath, objGroup, objUser, GALQueryFilter, strQuery

'Create a RootDSE Object, this will be used to connect to the domain
Set objRootDSE = GetObject("LDAP://rootDSE")

'Exit the script if a Domain isn't located
If Err Then
   MsgBox "You are not connected to a Domain.",vbCritical,"Domain Not Found"
   Err.Clear
   WScript.Quit
End If

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

'Loop Thru each member of the record set
While Not objRecordset.EOF
   If uCase(objRecordset.Fields("CN")) = "DOMAIN ADMINS" Then
      
      'Build the Group path and create the Group object
      strADSPath = "LDAP://" & objRecordset.Fields("DistinguishedName")
      Set objGroup = GetObject(strADSPath)
      
      'Loop thru each user in the group and find that administrator
      For Each objUser in objGroup.Members
         If UCase(objUser.msExchHomeServerName) <> "" Then
            strMailServerAddress = objUser.msExchHomeServerName
            Exit For
         End If         
      Next
   End If
   
   'Move to the next record
   objRecordset.MoveNext
Wend

'Exit the script if it couldn't find the mail server address.
If strMailServerAddress = "" Then
   MsgBox "The script could not determine the address of the mail server.",vbCritical,"Error - Mail Server Not Found"
   Wscript.Quit
End If

'Build the query string
GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & strMailServerAddress & ")) )))))"
strQuery = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & ">;" & GALQueryFilter & ";samaccountname;subtree"

'Create the connection object and set it's properties
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOOBJECT"
objConnection.Open "ADs Provider"

'Create the command object and set it's properties
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.CommandText = strQuery
objCommand.Properties("Sort on") = "givenname"

'Execute the command
Set objRecordset = objCommand.Execute

'Display the number of mailboxes
Msgbox "There are " & objRecordSet.RecordCount & " mailboxes on your mail server.",vbOkOnly,"Number of Mailboxes"