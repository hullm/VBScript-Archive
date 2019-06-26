'Created by Matthew Hull on 5/23/04
'Last Updated 2/07/05

'Version 1.01

'This script is designed to use the SubinACL program found in the Windows
'2000 Resource Kit to reset Ownership on users home folders.  It will 
'assign ownership of all files and folders in each users folder to the
'proper account. 

'Version History
'~~~~~~~~~~~~~~~
'Version 1.01 - Permissions will now be reset on parent folders as well as
'               subfolders and files.

'Version 1.0  - First version of this script released.

Option Explicit

On Error Resume Next

Dim strPath, objFSO, objShell, objNet, objfolder, colFolders, strRootPath
Dim strCMD, strDomain, intSubinACLError, strError, objPath

'**************************************************************************
'strPath = Source folder, UNC format will work
strPath = "D:\Shared Data\TeacherHome"
'**************************************************************************
'Do Not Change Anything Below This Line, All Data Entry is Above.

'Create the Network, File System and Shell objects.  The Network object will be
'used to retrieve domain information.  The File System object will be used to 
'create a collection object that contains all the folders you wish to modify.
'The Shell object is used to run the CACLS command.
Set objNet = WScript.CreateObject("WScript.Network")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("Wscript.Shell")

'This verifies that the root folder exists, if not the script will exit.
If Not objFSO.FolderExists(strPath) Then
   MsgBox "The folder " & strPath & " doesn't exist.  This script will now exit."
   Wscript.Quit   
End If

'Using the File System object to create a folder object, then create a collection
'object to store all the subfolders.
Set objPath = objFSO.GetFolder(strPath)
Set colFolders = objPath.Subfolders

'Use the Network object to get the domain name.
strDomain = objNet.UserDomain

'strRootPath is used to set the path during the loop.  It will add a "\" to the end if 
'it isn't already there.
If Right(strPath,1) = "\" Then
   strRootPath = strPath
Else
   strRootPath = strPath & "\"
End If

'This is where the SubinAC command is built and run.  If there is an error encountered
'the name of the folder will be added to strError.
For Each objFolder in colFolders
   strPath = strRootPath &  objFSO.GetBaseName(objFolder) & "\*.*"
   strCMD = "cmd /c subinacl /subdirectories " & """" & strPath & """ /SetOwner=" & strDomain & _
      "\" & objFolder.Name
   intSubinACLError = objShell.Run(strCMD,0,true) 
   strCMD = "cmd /c subinacl /file=directoriesonly " & """" & strRootPath &  objFSO.GetBaseName(objFolder) & """ /SetOwner=" & strDomain & _
      "\" & objFolder.Name
   intSubinACLError = objShell.Run(strCMD,0,true) 
   If intSubinACLError <> 0 Then
      MsgBox "SubinACL doesn't appear to be installed or it isn't located in a folder in the path", _
         vbCritical,"SubinACL Not Found"
      Wscript.Quit
   End If
Next

'Display a message when done
MsgBox "The folders have been modified in " & strRootPath,vbOkOnly,"Complete"   

'Close all open objects
Set objNet = Nothing
Set objFSO = Nothing
Set objShell = Nothing
Set objFolder = Nothing
Set colFolders = Nothing