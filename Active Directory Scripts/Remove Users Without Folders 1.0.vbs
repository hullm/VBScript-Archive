'Created by Matthew Hull on 9/1/05 

'This script will delete any user that doesn't have a home folder.  The idea 
'is that you would have run the script "Delete Empty Folders 1.0.vbs" first.
'This way anyone who has never used their account can be removed.
'USE AT YOUR OWN RISK - THIS IS DESIGNED TO DELETE USER ACCOUNTS!!!!

Option Explicit

On Error Resume Next

Dim strPath, strSourceOU, objADHelper, objFSO, objOU, objUser

'*************************************************************************************
'Enter the path to the home folders and the OU that contains the users.
strPath = "C:\Folder"
strSourceOU = "Students"
'*************************************************************************************

'Create a File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'See if the root folder exists, if not exit the script
If Not objFSO.FolderExists(strPath) Then
   MsgBox "The folder " & strPath & " doesn't exist.  This script will now exit."
   Wscript.Quit   
End If

'Used to get the OU object
Set objADHelper = CreateObject("ADHelper.wsc")

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
   "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Use the ADHelper object to get the OU object
Set objOU = objADHelper.OUObject(strSourceOU)

'Loop thru each user in the OU
For Each objUser in objOU

   'Verify the object is of type user
   If objUser.Class = "user" Then
      If Not objFSO.FolderExists(strPath & "\" & objUser.SamAccountName) Then
         objUser.DeleteObject(0)
      End If
   End If
Next

'Display a message to the user when done
MsgBox "Accounts have been removed.",vbOkOnly,"Complete"

'Close all open objects
Set objFSO = Nothing
Set objADHelper = Nothing