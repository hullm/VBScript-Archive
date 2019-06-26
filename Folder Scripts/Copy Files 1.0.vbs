'Created By Matthew Hull on some day a long time ago
'Documented on 4/25/04
'Last Updated 3/2/05

'This script will copy files from one folder to another.  If there 
'is a error it will let the user know what folder it failed on.

Option Explicit

On Error Resume Next

Dim objFSO, strSourceFolder, strDestFolder, colFolders, objFolder
Dim objRootFolder, strCopyError, WSHShell, colFiles, objFile

'This sub will copy files, it is missing an On Error Resume Next
'statement, this way it will exit if it fails.  If you wanted
'it to continue on an Access Denied error you could add it.
Sub CopyFiles

   'Create a folder object that is the root source folder
   Set objRootFolder = objFSO.GetFolder(strSourceFolder)
   
   'Create a collection of folders that contain the subfolders of
   'the root folder
   Set colFolders = objRootFolder.SubFolders

   'Loop thru each folder in the collection and copy it to the
   'destination
   For Each objFolder in colFolders
      objFolder.Copy(strDestFolder & "\" & objFolder.Name)
   Next

   Set colFiles = objRootFolder.Files

   'Loop thru each file in the collection and copy it to the
   'destination
   For Each objfile in colFiles
      objFile.Copy(strDestFolder & "\")
   Next
End Sub

'Create the File System Object, this will be used to create the
'folder objects later
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create a shell object that will be used to get a Special Folder
Set WSHShell = WScript.CreateObject("WScript.Shell")

'********************************************************************
'Enter the source and destination folders

strSourceFolder = "\\equake\c$\webupload"
strDestFolder = "D:\Website\AS1"
CopyFiles

'********************************************************************

'If an error is encoutered build the error message
If Err Then
   Err.Clear
   strCopyError = strCopyError & objFolder.Path & vbCRLF         
End If

'Let the user know the script is done and the status
If strCopyError = "" Then
   'MsgBox "Copy Complete",vbOkOnly,"Complete"
Else
   'MsgBox "You do not have access to copy the following folder:" & _
   'vbCRLF & vbCRLF & strCopyError & vbCRLF & "Please correct the " & _ 
   '"problem and rerun this script.",vbCritical,"Access Denied"
End If