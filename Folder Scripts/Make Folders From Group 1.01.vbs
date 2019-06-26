'Created by Matthew Hull on 10/11/03
'Last Updated 12/21/03
'Documented on 4/26/04

'Version 1.01

'This script will create a folder for every user in a security group.  It will also setup
'permissions on the folder.  If a folder for user already exists it will skip over it.

'Version History
'~~~~~~~~~~~~~~~
'Version 1.01 - Added strOwnerPermission to allow you to determine the owners permissions on
'               the folders that are created.
'             - Added an AD Search to locate the group by name.  You no longer have to type
'               in the LDAP location of the group.
'             - Added a dictionary object to allow you to add multiple users and groups to the
'               ACL of the folders created.
'             - Removed the need for a "/" at the end of strPath, and the need for a ","
'               at the end of the strSourceGroup.

'Version 1.0 - First version of this script released.

'*************************************************************************************

Option Explicit

On Error Resume Next

Dim strPath, strUserName, strCMD, strDomain, strADSPath, strSourceGroup
Dim objNet, objFSO, objShell, objGroup, objRootDSE, intCaclsError, objUser
Dim objPerms, strOwnerPermission, bolGroupFound, objConnection, objCommand
Dim objRecordSet, Key, strRootPath, intSuccess, intFail, intTotal

'This dictionary object will contain all user/groups and permissions to add
'to the new folders
Set objPerms = CreateObject("Scripting.Dictionary")

'**************************************************************************
'strPath = The location where all the folders will be created
'strSourceGroup = The name of the group where the users are located.
'strOwnerPermission = The permissions assigned to the owner of the folder.
'   - r=Read Only; w=Write; c=Change; f=Full
'**************************************************************************

strPath = "\\Server\Share"
strSourceGroup = "Security Group"
strOwnerPermission = "c"

'**************************************************************************
objPerms.Add "Domain Admins", "f"

'**************************************************************************

'Create a Net Object, this will be used to get the domain name
Set objNet = CreateObject("WScript.Network")

'Create a File System Object, This will be used to create the folders
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create a Shell Object, this will be used to run the CACLS command
Set objShell = CreateObject("Wscript.Shell")

'Create a RootDSE Object, this will be used to connect to the domain
Set objRootDSE = GetObject("LDAP://rootDSE")

'This will turn true if the user entered a valid group name
bolGroupFound = False

'Initalize counters
intSuccess = 0
intFail = 0
intTotal = 0

'See if the root folder exists, if not exit the script
If Not objFSO.FolderExists(strPath) Then
   MsgBox "The folder " & strPath & " doesn't exist.  This script will now exit."
   Wscript.Quit   
End If

'I don't know why but this script doesn't work on the Domain Users group,
'this will let the user know that if they try to do it.
If uCase(strSourceGroup) = "DOMAIN USERS" Then
   MsgBox "You cannot create folders for members of the Domain Users group.", _
   vbOkOnly,"Domain USers"
   Wscript.Quit
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

'Set the Root Path and get the domain name
strRootPath = strPath
strDomain = objNet.UserDomain

'Loop Thru each record in the record set.
While Not objRecordset.EOF
   If uCase(objRecordset.Fields("CN")) = uCase(strSourceGroup) Then
      
      'The group was found so this is changed to true
      bolGroupFound = True
      
      'Build the Group path and create the Group object
      strADSPath = "LDAP://" & objRecordset.Fields("DistinguishedName")
      Set objGroup = GetObject(strADSPath)
      
      'Loop thru each user in the group
      For Each objUser in objGroup.Members
      
         'Get the username
         strUserName = objUser.SAMAccountName
         
         'Build the users home folder path
         strPath = strRootPath & "\" & strUsername
         
         'Increase the total number of users by 1
         intTotal = intTotal + 1
         
         'If the user doesn't have a folder create one and set permissions
         If Not objFSO.FolderExists(strPath) Then
         
            'Increase the number of folders created by 1
            intSuccess = intSuccess + 1
            
            'Create the folder
            objFSO.CreateFolder(strPath)
            
            'Add the username to the Dictionary Object with their permission   
            objPerms.Add strUserName, strOwnerPermission
            
            'Build the CACLS command using the Dictionary Object
            strCMD = "cmd /c echo y| cacls " & """" & strPath & """ /c /t /g "
            For Each Key in objPerms
               strCMD = strCMD & """" & strDomain & "\" & Key & """" & ":" & objPerms.Item(Key) & " "      
            Next
            
            'Run the CALCS command and return any erroros to intCaclsError
            intCaclsError = objShell.Run(strCMD,0,true)
            
            'If there was an error during the CACLS command an error will be displaied
            'and the script will exit
            If intCaclsError <> 0 Then
               objFSO.DeleteFolder(strPath)
               MsgBox "There is a problem with a permissions setting..." & _
               vbCRLF & "No folders have been created.",vbOkOnly,"Permissions Error"
               Wscript.Quit
            End If
            objPerms.Remove(strUsername)
         Else
         
            'If the user already has a folder increase the fail count by one
            intFail = intFail + 1
         End If
      Next
   End If
   
   'Move to the next record
   objRecordset.MoveNext
Wend

'Close the connection with Active Directory
objConnection.Close

'If the group isn't located in Active Directory the script will exit
If Not bolGroupFound Then
   MsgBox "There is a problem with your strSourceGroup setting..." & _
   vbCRLF & "No folders have been created." & vbCRLF & _
   "strSourceGroup = " & strSourceGroup,vbOkOnly,"strSourceGroup Error"
   WScript.Quit
End If

'Display a message to the user when done
MsgBox intSuccess & " folders have been created in " & strRootPath & vbCRLF & _
intFail & " users already had folders." & vbCRLF & "There is a total of " & intTotal & _
" users in " & strSourceGroup & ".",vbOkOnly,"Complete"

'Close objects
Set objNet = Nothing
Set objFSO = Nothing
Set objShell = Nothing
Set objRootDSE = Nothing
Set objGroup = Nothing
Set objPerms = Nothing
Set objCommand = Nothing
Set objRootDSE = Nothing
Set objRecordSet = Nothing