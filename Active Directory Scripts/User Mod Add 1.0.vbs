'Created by Matthew Hull on 8/8/04
'Last Updated 8/9/04

'This script will modify users properties in Active Directory

Option Explicit

On Error Resume Next

Dim strOU, strDescription, strHomeDirectory, strHomeDrive, strProfilePath, strScriptPath
Dim objADHelper, objOU, objUser, bolUseUsername, bolSharedProfile, strHomeRoot
Dim strProfileRoot

'***************************************************************************************
'All user entry is done in this section.  Leave the value blank, ie "", if you don't
'want to modify that property.

strOU = "ScriptTest"

strDescription = "Test"

strHomeDirectory = "\\Server\Share"
   bolUseUsername = True 'Set this to true to have each user use their own folder
strHomeDrive = "Z:"

strProfilePath = ""
   bolSharedProfile = True 'Set this to true if all users point to the same profile

strScriptPath = "Script.bat"
'***************************************************************************************

'Create the ADHelper object
Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
   "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Create an OU object using the ADHelper object
Set objOU = objADHelper.OUObject(strOU)

'Exit the script if the OU isn't found
If Err Then
   MsgBox """" & strOU & """ is not a valid OU.  The script will now exit.", _
   vbCritical,"Invalid OU"
   WScript.Quit
End If

'Set the root folder to be used to set the property
If bolUseUsername Then
   If Right(strHomeDirectory ,1) <> "\" Then
      strHomeRoot = strHomeDirectory & "\"
   Else
      strHomeRoot = strHomeDirecotry
   End If
End If

'Set the root folder to be used to set the property   
If Not bolSharedProfile Then
   If Right(strProfilePath ,1) <> "\" Then
      strProfileRoot = strProfilePath & "\"
   Else
      strProfileRoot = strProfilePath
   End If
End If
   
'Loop thru each user in the OU
For Each objUser in objOU

   'Verify the user is of type user
   If objUser.Class = "user" Then
      
      If strDescription <> "" Then
         objUser.Put "description", strDescription
      End If
      
      If strHomeDirectory <> "" Then
         If bolUseUsername Then
            strHomeDirectory = strHomeRoot & objUser.SamAccountName
         End If
         objUser.Put "homedirectory", strHomeDirectory
      End If
      
      If strHomeDrive <> "" Then
         objUser.Put "homedrive", strHomeDrive
      End If
      
      If strProfilePath <> "" Then
         If Not bolSharedProfile Then
            strProfilePath = strProfileRoot & objUser.SamAccountName   
         End If   
         objUser.Put "profilepath", strProfilePath
      End If
      
      If strScriptPath <> "" Then
         objUser.Put "scriptpath", strScriptPath
      End If
      
      objUser.SetInfo      
   End If
Next

'Display a message when complete
MsgBox "OU Modifications Complete",vbOkOnly,"Complete"

'Close open objects
Set objADHelper = Nothing
Set objOU = Nothing
Set objUser = Nothing