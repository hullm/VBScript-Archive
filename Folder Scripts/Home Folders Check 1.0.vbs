'Created by Matthew Hull a long time ago
'Last Updated 5/23/04
'Documented on 4/24/04

'This script was written for a school to see if users have home folders

Option Explicit

On Error Resume Next

Dim objFSO, objOU, strUser, strRootPath, strPath, intCount, strMessage

'************************************************************************************

'Enter the home folder locaton, and the OU you want to scan
strRootPath = "\\server\share"
strOU = "LDAP://OU=Students,OU=STUDENTS,OU=ACCOUNTS,DC=Domain,DC=ORG"

'************************************************************************************

'Create the File System Object used to see if the users have folders
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create the OU object using the input from earlier
Set objOU = GetObject(strOU)

'Start the count of missing folders at 0
intCount = 0

'Loop Thru each user in the OU.  NOTE: At this point we have On Error Resume Next
'turned on.  If the user entered the OU string wrong there would have been an error
'generated during the creation of the OU object.  There should be an error check
'at this point
For Each strUser in objOU
   
   'Build the home folder path using the users username
   strPath = strRootPath & strUser.SAMAccountName

   'See if the user has a home folder
   If Not objFSO.FolderExists(strPath) Then

      'Incresase the count by one for each user without a folder
      intCount = intCount + 1
      
      'Add the user to the message
      strMessage = strMessage & intCount & " " & strUser.SamAccountName & vbCRLF
   End If 
Next

'Display a message to the user when done
MsgBox strMessage

'Close all objects
Set objFSO = Nothing
Set objOU = Nothing