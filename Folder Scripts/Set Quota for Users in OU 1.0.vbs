'Created by Matthew Hull on 8/31/05

'This script will set a quota limit for everyone in an OU.
'NOTE: This version works but can use some improvements, like an exception list
'and it could easly be made to run on multiple OU's at once.  Maybe with the next
'version.

Option Explicit

On Error Resume Next

Dim objADHelper, strOU, objOU, objUser, strThreshold, strLimit, strCMD, intCMDError

'*************************************************************************************
'Set quota's on the following drive.
strDrive = "d:"

'Enter the OU name then set the Warning Threshold and Limit, both in MB's
strOU = "Students"
strThreshold = "10"
strLimit = "7"

'*************************************************************************************

'Used to get the OU object
Set objADHelper = CreateObject("ADHelper.wsc")

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
   "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Create a Shell Object, this will be used to run the fsutil command
Set objShell = CreateObject("Wscript.Shell")

SetLimits

Sub SetLimits

   strLimit = strLimit * 1048576
   strThreshold = strThreshold * 1048576
   
   Set objOU = objADHelper.OUObject(strOU)
   
   For Each objUser in objOU
      If objUser.Class = "user" Then
         strCMD = "cmd /c fsutil quota modify " & strDrive & " " & strThreshold & " " & strLimit & " " & objUser.SamAccountName
         intCMDError = objShell.Run(strCMD,0,true)
      End If
   Next
   
   strLimit = strLimit \ 1048576
   strThreshold = strThreshold \ 1048576
   
End Sub

'Display a message when complete
MsgBox "Complete",vbOkOnly,"Done"

'Close Open Objects
Set objOU = Nothing
Set objShell = Nothing
Set objUser = Nothing
Set objADHelper = Nothing