'Created by Matthew Hull on 10/17/03
'Documented on 4/24/04
'Last Updated 8/9/04

'This script will delete all profiles on a local computer besides
'Default User, All Users, Administrator, and the current logged on
'user.  The portion that deletes the folder is commented out.

Option Explicit

On Error Resume Next

'Declare variables
Dim objFSO, colDrives, colFolders, objDelFolder, objDrives, strConfirm
Dim strFolder, strRootFolder, objFolder, strPath, intCount, strDrive
Dim strUserName, objWSHShell, strBeforeSize, strAfterSize, strSize
Dim strProfilePath

'Create the Shell object to grab environmental variables from OS
Set objWSHShell = CreateObject("WScript.Shell")

'Verifies that the user is running an NT based OS
If objWSHShell.ExpandEnvironmentStrings("%OS%") <> "Windows_NT" Then
   MsgBox "This program will not run on Windows 9x/ME.", _
   vbOkOnly,"Wrong OS"
   WScript.Quit
End If

'Verify that the user wants to delete folders.  Converts the result 
'to capital letters.  If the result isn't YES then the scripts exits.
strConfirm = InputBox("This script will delete all users profiles on " & _
"this computer.  To continue type ""YES""","Confirm File Deletion")
strConfirm = UCase(strConfirm)
If strConfirm <> "YES" Then
   WsCript.Quit
End If

'Create the file system object that will do the folder deletion
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create a collection of dirves, this will be used to get the size before
'the folder deletion, then get the size after.  
Set colDrives = objFSO.Drives

'Get the currect users username, we don't want to try to delete the current
'users profile
strUserName = uCase(objWSHShell.ExpandEnvironmentStrings("%username%"))

'Get the All Users Profile path, then strip of the end to get the profile path
strProfilePath = objWSHShell.ExpandEnvironmentStrings("%AllUsersProfile%")
strRootFolder = (objFSO.GetParentFolderName(strProfilePath)) & "\"

'Check to see if the profile path exists, if not exit the script
If Not objFSO.FolderExists(strRootFolder) Then
   Wscript.Quit   
End If

'Loop thru each drive in the collection of drives and get the first letter
'then compair it to the first letter of the profile path.  If they are the 
'same get the size of that drive.  This will allow us to know how much space
'has been cleaned
For Each strDrive in colDrives
   If (Left(strDrive.DriveLetter,1)) = (Left(strProfilePath,1)) Then
      strBeforeSize = strDrive.FreeSpace
   End If
Next

'This variable will hold the number of deleted folders
intCount = 0

'Create a folder object that is the profile path then create a collection
'of folders that contain all the subfolders of the profile path.  These 
'are the folders we want to delete
Set objFolder = objFSO.GetFolder(strRootFolder)
Set colFolders = objFolder.SubFolders

'Loop thru each folder in the collection
For Each strFolder in colFolders

   'Check the name of the folder, if it is not one of the reserved names
   'then detete it.
   Select Case UCase(strFolder.Name)
      Case "DEFAULT USER"
      Case "ALL USERS"
      Case "ADMINISTRATOR"
      Case "NETWORKSERVICE"
      Case "LOCALSERVICE"
      Case strUserName
      Case Else
      
         'Build the path the the folder that needs to be deleted
         strPath = strRootFolder & strFolder.Name
         
         'Create a folder object for the folder that needs to be deleted
         Set objDelFolder = objFSO.GetFolder(strPath) 
        
'**********************************************************
        'The next line deletes the folders. Uncomment it.
        'objDelFolder.Delete(true)
'**********************************************************

         'If the deletion was a success then increase the count
         If Err Then
            Err.Clear            
         Else 
            intCount = intCount + 1
         End If
         
      End Select
Next

'Loop thru each drive in the collection of drives and get the first letter
'then compair it to the first letter of the profile path.  If they are the 
'same get the size of that drive.  This will allow us to know how much space
'has been cleaned
For Each strDrive in colDrives
   If (Left(strDrive.DriveLetter,1)) = (Left(strProfilePath,1)) Then
      strAfterSize = strDrive.FreeSpace
   End If
Next

'Calculate the amount of data that has been deleted
strSize = Round(((strAfterSize - strBeforeSize)/1048576),2)

'If the size is 0 then display a message about no folder deleted, if not then
'display the number of folders deleted and the amount of space freed
If strSize = 0 Then
   MsgBox "Access Denied or you need to uncomment the line that " & _
   "deletes data or there is nothing to delete.", _
   vbOkOnly,"Access Denied"
Else
   MsgBox intCount & " folders have been deleted.  You freed " & _
   strSize & " Mbytes on your computer.",vbOkOnly,"Complete"
End If

'Close objects
Set objFSO = Nothing
Set objWSHShell = Nothing
Set objFolder = Nothing