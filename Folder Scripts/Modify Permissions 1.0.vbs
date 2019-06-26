'Created by Matthew Hull on 11/09/03
'Last Updated 12/27/03

'Version 1.0

'This script will reset permissions on all files and folders in a source folder.
'If bolUseUserName is set to true it will add the user with the same username as the
'folder name to the ACL of that folder.  If there are errors during the process it
'will log which folders failed.  Most of the time it fails because no user exists that
'has a username that matches the folders name.  This can be used to reset permissions
'on users home folders.  Note: You may have to take ownership first to have rights
'to run this script.

'**************************************************************************

Option Explicit

On Error Resume Next

Dim strPath, objFSO, objShell, objNet, objfolder, colFolders, strRootPath
Dim strCMD, strDomain, bolUseUserName, intCaclsError, objDict, Key
Dim strError, bolError, intCount, bolRePlacePermissions, strReplace
Dim strErrorOutput, txtOutPut, objPath, strOwnerPermission

Set objDict = CreateObject("Scripting.Dictionary")

'**************************************************************************
'bolUseUserName = True or False if True the user will be added to the ACL
'   StrOwnerPermission - r=Read Only; w=Write; c=Change; f=Full

'bolReplacePermissions = True or False if True the ACL's will be replaced
'strPath = Source folder, UNC format will work
'strErrorOutput = Location where error log will be stored
'**************************************************************************

bolUseUserName = True
   strOwnerPermission = "c"

bolReplacePermissions = True
strPath = "C:\Test"
strErrorOutput = "C:\error.txt"

'**************************************************************************
'objDict.Add "Group Name", "Permission"
'   - r=Read Only; w=Write; c=Change; f=Full
'**************************************************************************

objDict.Add "Domain Admins", "f"

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

'This will turn on the /e switch in the CACLS command which will set it to edit
'the ACL instead of replacing it.
If bolReplacePermissions Then
   strReplace = ""
Else
   strReplace = "/e "
End If

'Using the File System object to create a folder object, then create a collection
'object to store all the subfolders.
Set objPath = objFSO.GetFolder(strPath)
Set colFolders = objPath.Subfolders

'Use the Network object to get the domain name.  intCount will store the number of 
'errors encountered while running CACLS.  strRootPath is used to set the path during
'the loop.  It will add a "\" to the end if it isn't already there.
strDomain = objNet.UserDomain
intCount = 0
If Right(strPath,1) = "\" Then
   strRootPath = strPath
Else
   strRootPath = strPath & "\"
End If

'This is where the CACLS command is built and run.  If there is an error encountered
'the name of the folder will be added to strError.
For Each objFolder in colFolders
   strPath = strRootPath &  objFSO.GetBaseName(objFolder)
   If bolUseUserName Then
      objDict.Add objFolder.Name, strOwnerPermission
   End If
   strCMD = "cmd /c echo y| cacls " & """" & strPath & """ /c /t " & strReplace & "/g "
   For Each Key in objDict
      strCMD = strCMD & """" & strDomain & "\" & Key & """" & ":" & objDict.Item(Key) & " "      
   Next
   intCaclsError = objShell.Run(strCMD,0,true) 
   If intCaclsError <> 0 Then
      intCOunt = intCount + 1
      strError = strError & vbCRLF & objFSO.GetBaseName(objFolder)
      bolError = True
   End If
   If bolUseUserName Then
      objDict.Remove(objFSO.GetBaseName(objFolder))
   End If
Next

'If there was an error this will display which folders failed and write the output to a
'file.  If all is ok it will let us know the script is done.
If bolError Then
   MsgBox "The folders have been modified.  There was a problem setting " & intCount & " folders up " & _
   "permissions on the following folders:" & strError,vbOkOnly,"Complete with Errors"
   If strErrorOutput <> "" Then
      Set txtOUtPut = objFSO.CreateTextFile(strErrorOutput)
      txtOutput.Write "The folders have been modified.  There was a problem setting " & intCount & _
      " folders up permissions on the following folders:" & strError
      txtOutput.close
   End If
Else
   MsgBox "The folders have been modified in " & strRootPath,vbOkOnly,"Complete"   
End If

'Close all open objects
Set objNet = Nothing
Set objFSO = Nothing
Set objShell = Nothing
Set objFolder = Nothing
Set colFolders = Nothing
Set objDict = Nothing