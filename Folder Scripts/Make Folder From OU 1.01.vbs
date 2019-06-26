'Created by Matthew Hull on 10/11/03 
'Last Updated 12/21/03
'Documented on 4/26/04

'Version 1.01

'This script will create a folder for every user in an OU.  It will also setup
'permissions on the folders.  If a folder for the user already exists it will
'skip over it.

'Version History
'~~~~~~~~~~~~~~~
'Version 1.01 - Added strOwnerPermission to allow you to determine  the owners permissions on
'               the folders that are created.
'             - Added an AD Search to locate the OU by name.  You no longer have to type
'               in the LDAP location of the OU.  If multiple OU's are found with the same
'               It will prompt you to shoose the correct OU.
'             - Added a dictionary object to allow you to add multiple users and groups to the
'               ACL of the folders created.
'             - Removed the need for a "/" at the end of strPath, and the need for a ","
'               at the end of the strSourceOU.

'Version 1.0 - First version of this script released.

'*************************************************************************************

Option Explicit

On Error Resume Next

Dim strPath, strUserName, objPerms, strCMD, strDomain, strADSPath, strSourceOU, Key
Dim objNet, objFSO, objShell, objOU, objRootDSE, intCaclsError, strUser, strRootPath
Dim strOwnerPermission, intSuccess, intFail, intTotal, intCount, strInput
Dim objConnection, objCommand, objRecordSet, objOUList, strDN, intChoice

'This dictionary object will contain all user/groups and permissions to add
'to the new folders
Set objPerms = CreateObject("Scripting.Dictionary")

'**************************************************************************
'strPath = The location where all the folders will be created
'strSourceOU = The location where the user accounts will be found.
'   - if more then one is located with the name it will prompt you
'strOwnerPermission = The permissions assigned to the owner of the folder.
'   - r=Read Only; w=Write; c=Change; f=Full
'**************************************************************************

strPath = "\\Server\Share"
strSourceOU = "Users"
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

'Establish a connection to Active Directory using ActiveX Data Object
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADSDSOObject;"

'Create the command object and attach it to the connection object
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

'objCommand defines the search base (in this case the whole domain) 
'a filter (All object of type OrganizationalUnit) and some attributes 
'associated with the returned objects (Name and DistinguishedName)
objCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
">;(&(objectClass=OrganizationalUnit));Name,DistinguishedName"

'Initiate the LDAP query and return results to a RecordSet object.
Set objRecordSet = objCommand.Execute

'Create a Dictionary Object that will hold all the OU's found
Set objOUList = CreateObject("Scripting.Dictionary")

'Start the OU count at 0
intCount = 0

'If more then one OU is found this will be the error message displayed.
strInput = "There was more then one OU found with that name, please "
strInput = strInput & "choose the correct OU from the list below"
strInput = strInput & vbCRLF
strInput = strInput & "The input will be required on the next screen." & vbCRLF

'Loop thru the returned RecordSet and look for the OU.  If found it adds the 
'DN name to objOUList and the error message. It will also increases intCount by 1
While Not objRecordset.EOF
   If uCase(objRecordset.Fields("Name")) = UCase(strSourceOU) Then
      intCount = intCount + 1
      strInput = strInput & intCount & " " & _
      objRecordset.Fields("DistinguishedName") & vbCRLF
      strDN = objRecordSet.Fields("DistinguishedName")
      objOUList.Add "Key" & intCount, strDN
   End If
   objRecordset.MoveNext
Wend

'Close the connection with Active Directory
objConnection.Close

'If more then one OU is found it will ask you which you want, if only one is
'found it will set intChoice to 1.  If no OU is found the script will exit.
If intCount > 1 Then
   Do
      MsgBox strInput,vbOKOnly,"More Then One OU Found"
      intChoice = InputBox("Please choose your OU","Select OU")
      If objOUList.Item("Key" & intChoice) = "" Then
         MsgBox "You must choose a valid OU to continue.",vbOkOnly,"Error"
      End If
   Loop Until objOUList.Item("Key" & intChoice) <> ""
ElseIf intCount = 0 Then
   MsgBox strOU & " is not a valid OU.",vbOkOnly,"Invalid OU"
   WScript.Quit
Else
   intChoice = 1
End If

'Build the OU path and create the OU object
strADSPath = "LDAP://" & objOUList.Item("Key" & intChoice)
Set objOU = GetObject(strADSPath)

'Initalize counters
intSuccess = 0
intFail = 0
intTotal = 0

'This will look for an error in the OU object creation, if found the script will
'exit.  This is left over from 1.0 and could be removed.  Since the domain is
'looking for OU for you the error checking is done before this
If Err Then
   MsgBox "There is a problem with your strSourceOU setting..." & _
   vbCRLF & "No folders have been created." & vbCRLF & _
   "strSourceOU = " & strSourceOU,vbOkOnly,"strSourceOU Error"
   Err.Clear
   WScript.Quit
End If

'See if the root folder exists, if not exit the script
If Not objFSO.FolderExists(strPath) Then
   MsgBox "The folder " & strPath & " doesn't exist.  This script will now exit."
   Wscript.Quit   
End If

'Set the Root Path and get the domain name
strRootPath = strPath
strDomain = objNet.UserDomain

'Loop thru each user in the OU
For Each strUser in objOU

   'Verify the object is of type user
   If strUser.Class = "user" Then
   
      'Get the username
      strUserName = strUser.SAMAccountName
      
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
   End If
Next

'Display a message to the user when done
MsgBox intSuccess & " folders have been created in " & strRootPath & vbCRLF & _
intFail & " users already had folders." & vbCRLF & "There is a total of " & intTotal & _
" users in OU " & strSourceOU & ".",vbOkOnly,"Complete"

'Close all open objects
Set objNet = Nothing
Set objFSO = Nothing
Set objShell = Nothing
Set objRootDSE = Nothing
Set objPerms = Nothing
Set objCommand = Nothing
Set objRootDSE = Nothing
Set objRecordSet = Nothing