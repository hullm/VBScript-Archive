'Created by Matthew Hull on 8/8/04

'This script will clear certain users properties in Active Directory

Option Explicit

On Error Resume Next

Const ADS_PROPERTY_CLEAR = 1

Dim strOU, bolDescription, bolHomeDirectory, bolHomeDrive, bolProfilePath, bolScriptPath
Dim objADHelper, objOU, objUser

'***************************************************************************************
'All user entry done in this section.  Set the value to true if you want to clear it

strOU = "OU Name"
bolDescription = False
bolHomeDirectory = False
bolHomeDrive = False
bolProfilePath = False
bolScriptPath = False
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

'Loop thru each user in the OU
For Each objUser in objOU

   'Verify the user is of type user
   If objUser.Class = "user" Then
      
      If bolDescription Then
         objuser.putEx ADS_PROPERTY_CLEAR, "description", 0
      End If
      
      If bolHomeDirectory Then
         objUser.PutEx ADS_PROPERTY_CLEAR, "homedirectory", 0
      End If
      
      If bolHomeDrive Then
         objUser.PutEx ADS_PROPERTY_CLEAR, "homedrive", 0
      End If
      
      If bolProfilePath Then
         objUser.PutEx ADS_PROPERTY_CLEAR, "profilepath", 0
      End If
      
      If bolScriptPath Then
         objUser.PutEx ADS_PROPERTY_CLEAR, "scriptpath", 0
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