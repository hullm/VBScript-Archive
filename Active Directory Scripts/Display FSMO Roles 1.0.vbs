'Created by Matthew Hull on 4/27/04

'This script will display the FSMO role holders for your domain

Option Explicit

On Error Resume Next

Dim objRootDSE, strDomainDN, strConfigDN, objPDCFSMO, strPDC, objPDCServer
Dim strMessage, objRIDFSMO, strRID, objRIDServer, objSchemaFSMO, strSchema
Dim objSchemaServer, objInfraFSMO, strInfra, objInfraServer, objDNFSMO
Dim strDN, objDNServer, strSchemaDN

'Create a RootDSE Object, this will be used to connect to the domain
Set objRootDSE = GetObject("LDAP://rootDSE")

'Exit the script if a Domain isn't located
If Err Then
   MsgBox "You are not connected to a Domain.",vbCritical,"Domain Not Found"
   Err.Clear
   WScript.Quit
End If

'Get the Default Naming Context
strDomainDN  = objRootDSE.Get("DefaultNamingContext")

'Get the Schema Naming Context
strSchemaDN = objRootDSE.Get("SchemaNamingContext")

'Get the Configuration Naming Context
strConfigDN = objRootDSE.Get("ConfigurationNamingContext")

'********************************************************************************

'Get the PDC Emulator and save it to a string
Set objPDCFsmo = GetObject("LDAP://" & strDomainDN)
strPDC = objPDCFSMO.FSMORoleOwner

'Connect to the server that is that contains the role
Set objPDCServer = GetObject("LDAP://" & Right(strPDC,Len(strPDC) - 17))

'Add the role to the message
strMessage = "PDC Emulator: " & objPDCServer.CN & vbCRLF

'********************************************************************************

'Get the RID Master and save it to a string
Set objRIDFSMO = GetObject("LDAP://cn=RID Manager$,cn=system," & strDomainDN)
strRID = objRIDFSMO.FSMORoleOwner

'Connect to the server that is that contains the role
Set objRIDServer = GetObject("LDAP://" & Right(strRID,Len(strRID) - 17))

'Add the role to the message
strMessage = strMessage & "RID Master: " & objRIDServer.CN & vbCRLF

'********************************************************************************

'Get the Schema Master and save it to a string
Set objSchemaFSMO = GetObject("LDAP://" & strSchemaDN)
strSchema = objSchemaFSMO.FSMORoleOwner

'Connect to the server that is that contains the role
Set objSchemaServer = GetObject("LDAP://" & Right(strSchema,Len(strSchema) - 17))

'Add the role to the message
strMessage = strMessage & "Schema Master: " & objSchemaServer.CN & vbCRLF

'********************************************************************************

'Get the Infrastructure Master and save it to a string
Set objInfraFSMO = GetObject("LDAP://cn=Infrastructure," & strDomainDN)
strInfra = objInfraFSMO.FSMORoleOwner

'Connect to the server that is that contains the role
Set objInfraServer = GetObject("LDAP://" & Right(strInfra,Len(strInfra) - 17))

'Add the role to the message
strMessage = strMessage & "Infrastructure Master: " & objInfraServer.CN & vbCRLF

'********************************************************************************

'Get the Domain Naming Master and save it to a string
Set objDNFSMO = GetObject("LDAP://cn=Partitions," & strConfigDN)
strDN = objDNFSMO.FSMORoleOwner

'Connect to the server that is that contains the role
Set objDNServer = GetObject("LDAP://" & Right(strDN,Len(strDN) - 17))

'Add the role to the message
strMessage = strMessage & "Domain Naming Master: " & objDNServer.CN

'********************************************************************************

'Display a message to the user when done
MsgBox strMessage,vbOkOnly,"FSMO Role Owners"

'Close Objects
Set objRootDSE = nothing
Set objPDCFsmo = nothing
Set objPDCServer = nothing
Set objRIDFSMO = nothing
Set objRIDServer = nothing
Set objSchemaFSMO = nothing
Set objSchemaServer = nothing
Set objInfraFSMO = nothing
Set objInfraServer = nothing
Set objDNFSMO = nothing
Set objDNServer = nothing