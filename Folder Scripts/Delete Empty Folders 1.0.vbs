'Created By Matthew Hull on 9/1/05

'This script will delete all the empty subfolders in a folder.  This is used to remove
'empty student folders. 
'USE AT YOUR OWN RISK - THIS IS DESIGNED TO DELETE DATA!!!!

Option Explicit

On Error Resume Next

Dim objFSO, colFolders, objRootFolder, strPath, objFolder

'*************************************************************************************
'Enter the path to the root folder
strPath = "c:\Folder"
'*************************************************************************************

'Create the File System Object, this will get the folder information
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Verify the folder exists, if not then exit the script
If Not objFSO.FolderExists(strPath) Then
   MsgBox "The folder " & strPath & " doesn't exist.",vbOkOnly,"Folder Doesn't Exist"
   Set objFSO = Nothing
   Wscript.Quit
End If

'Create the folder object
Set objRootFolder = objFSO.GetFolder(strPath)

'Create a collection of folder objects that contain the subfolders
Set colFolders = objRootFolder.SubFolders

'Write the report by looping thru each folder object in the collection
For Each objFolder in colFolders
   If objFolder.Size = 0 Then
      objFolder.Delete(True)
   End If
Next

'Display a message to the user when completet
MsgBox "Empty Folders Removed...",vbOkOnly,"Complete"

'Close objects
Set objFSO = Nothing
Set objRootFolder = Nothing
Set colFolders = Nothing