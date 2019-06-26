'Created By Matthew Hull on 7/28/04

'This script swaps the first and last name fields for all users in an OU.

Option Explicit

On Error Resume Next

Dim strOU, objADHelper, objOU, objUser, strFirstName, strLastName

'***************************************************************************************
strOU = "ScriptTest"
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
      strFirstName = objUser.SN
      strLastName = objUser.GivenName
      objUser.Put "GivenName", strFirstName
      objUser.Put "SN", strLastName
      objUser.SetInfo
   End If
Next

'Display a message when done
MsgBox "The users first and last name's have been swapped",vbOkOnly,"Done"

'Close open objects
Set objADHelper = Nothing
Set objOU = Nothing
Set objUser = Nothing