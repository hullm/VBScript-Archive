'Created By Matthew Hull on 7/26/04
'Last Updated 8/12/04

'Version 1.01

'This script will move users home folders from one location to another
'for all users in an OU.  It will update the profile and fix permissions.

'Version History
'~~~~~~~~~~~~~~~
'Version 1.01 - The script will modify the home folder settings for each user.
'             - Permissions will be set after the folder has been moved or copied.
'             - An error log will be generated with the error description for each
'               failure

'Version 1.0  - First version of this script released.

Option Explicit

On Error Resume Next

Dim objFSO, objOU, objUser, strRootPath, strPath, intCount, strDest
Dim strError, objADHelper, strOU, bolMove, strType, strErrorOutput
Dim objNet, objDict, objShell, strOwnerPermission, strDomain
Dim strCMD, Key, intErrorCount, txtOutPut

Set objDict = CreateObject("Scripting.Dictionary")

'******************************************************************************************
'All user entry is done in this section.

strOU = "ScriptTest" 'Name of the OU that contains the users
strRootPath = "\\Starbase\test" 'Source root folder
strDest = "\\Starbase\test2" 'Destination - Use a UNC Path
bolMove = False 'Set to true if you want to move the folders, set to False to copy
strErrorOutput = "C:\Move HF Error Log.txt"

'Set the permissions on the new folders by adding each group to the dictionary object
'objDict.Add "Group Name", "Permission"
'   - r=Read Only; w=Write; c=Change; f=Full

objDict.Add "Domain Admins", "f"

strOwnerPermission = "c" 'Set the permissions for the owner
'******************************************************************************************

'Create the required objects.
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objADHelper = CreateObject("ADHelper.wsc")
Set objNet = CreateObject("WScript.Network")
Set objShell = CreateObject("Wscript.Shell")

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
      "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Create the OU object
Set objOU = objADHelper.OUObject(strOU)

'Exit if the OU isn't found
If Err Then
   MsgBox "The OU " & strOU & " was not found.  The script will now exit.", _
   vbCritical,"OU Not Found"
   Err.Clear
   WScript.Quit
End If

'Fix strRootPath
If Right(strRootPath,1) <> "\" Then
   strRootPath = strRootPath & "\"
End If

'Fix strSource
If Right(strDest,1) <> "\" Then
   strDest = strDest & "\"
End If

'Initalize counters and get the domain name
intCount = 0
intErrorCount = 0
strDomain = objNet.UserDomain

'Loop thru each user in the OU
For Each objUser in objOU
   
   'Verify the user is of type user
   If objUser.Class = "user" Then

      'Build the users home folder path
      strPath = strRootPath & objUser.SAMAccountName
      
      'Change the users home folder path
      objUser.Put "homedirectory", strDest & objUser.SAMAccountName
      objUser.SetInfo
      
      'See if the user has a folder, if so move it and increase the count
      'if an error is encounterd then add the folder name to the error message   
      If objFSO.FolderExists(strPath) Then
         objFSO.CopyFolder strPath, strDest
         
         'If there is an error attempt to take ownership and reset permissions
         If Err Then
            strCMD = "cmd /c subinacl /subdirectories " & """" & strPath & """ /SetOwner=" & _
            strDomain & "\Administrator"
            Call objShell.Run(strCMD,0,true)
            
            strCMD = "cmd /c echo y| cacls " & """" & strPath & """ /c /t /g " & """" & _ 
            strDomain & "\Domain Admins""" & ":f" 
            Call objShell.Run(strCMD,0,true)
            Err.Clear
            
            objFSO.CopyFolder strPath, strDest
         End If
                  
         If Err Then
            intErrorCount = intErrorCount + 1
            strError = strError & intErrorCount & " - " & objUser.SAMAccountName & _
            " - Reason: " & Err.Description & vbCRLF
            Err.Clear
         Else
            intCount = intCount + 1
            If bolMove Then
               objFSO.DeleteFolder strPath, True
               strType = "moved"
            Else
               strType = "copied"
            End If
            
            'Set the permissions on the destination folder
            objDict.Add objUser.SamAccountName, strOwnerPermission
            strCMD = "cmd /c echo y| cacls " & """" & strDest & objUser.SAMAccountName & """ /c /t /g "
            For Each Key in objDict
               strCMD = strCMD & """" & strDomain & "\" & Key & """" & ":" & objDict.Item(Key) & " "      
            Next
            Call objShell.Run(strCMD,0,true)
            objDict.Remove(objUser.SamAccountName)
            
         End If
      End If
   End If 
Next

'Display a message when done
If strError = "" Then
   MsgBox intCount & " folders that have been " & strType & ".", _
   vbOkOnly,"Complete"
Else
   MsgBox intCount & " folders that have been " & strType & "." & vbCRLF & _
   "There was an error with the following folders: " & vbCRLF & strError & _
   "This message can be found in " & strErrorOutput, _
   vbOkOnly,"Complete with Errors"
   
   'Write any errors to the log file
   Set txtOutPut = objFSO.CreateTextFile(strErrorOutput)
   txtOutput.Write "The following folders did not get " & strType & "." & vbCRLF & _
   vbCRLF & strError
   txtOutput.Close
   Set txtOUtPut = Nothing   
End If

'Close objects
Set objFSO = Nothing
Set objOU = Nothing
Set objADHelper = Nothing
Set objUser = Nothing
Set objNet = Nothing
Set objDict = Nothing
Set objShell = Nothing