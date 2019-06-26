'Created by Matthew Hull on 7/8/04
'Last Updated 7/22/04

'This will Set e-mail storage limits.

Option Explicit

On Error Resume Next

'***********************************************************************************************
'Start of script, Declairing objects and variables
'***********************************************************************************************

Dim strOU, strWarning, strStopSend, strStopSendRecieve, bolEnableSettings

'***********************************************************************************************
'Main Body, - Set variable then call sub
'***********************************************************************************************

StrOU = "Script Test"
strWarning = "75000"
strStopSend = "100000"
strStopSendRecieve = "150000"
bolEnableSettings = True
FixMailboxLimits

MsgBox "The users have been modified.",vbOkOnly,"Done"

'***********************************************************************************************
'Fix Address Boox Sub, This is where the work gets done
'***********************************************************************************************

Sub FixMailboxLimits

   Const ADS_PROPERTY_CLEAR = 1

   Dim objUser, objOU, objADHelper
   
   On Error Resume Next   
   
   'Create ADHelper object
   Set objADHelper = CreateObject("ADHelper.wsc")
   
   'Exit if the ADHelper object isn't installed
   If Err Then
      MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
         "The script will now exit.",vbCritical,"Network Missing"
      Err.Clear
      WScript.Quit
   End If
   
   'Create OU object
   Set objOU = objADHelper.OUObject(strOU)
   
   'Loop thru each user in the OU and give them the site name
   For Each objUser in objOU
      If objUser.Class = "user" Then
      
         If strWarning = "" Then
            objUser.PutEx ADS_PROPERTY_CLEAR, "MDBStorageQuota", 0
         Else
            objUser.Put "MDBStorageQuota", strWarning
         End If
         
         If strStopSend = "" Then
            objUser.PutEx ADS_PROPERTY_CLEAR, "MDBOverQuotaLimit", 0
         Else
            objUser.Put "MDBOverQuotaLimit", strStopSend
         End If
         
         If strStopSendRecieve = "" Then
            objUser.PutEx ADS_PROPERTY_CLEAR, "MDBOverHardQuotaLimit", 0
         Else
            objUser.Put "MDBOverHardQuotaLimit", strStopSendRecieve
         End If
         
         If bolEnableSettings Then
            objUser.Put "MDBUseDefaults", False
         Else
            objUser.Put "MDBUseDefaults", True
         End If
         
         objUser.SetInfo
      End If
   Next
   
   'Exit if the user fails to update
   If Err Then
      MsgBox "Can not set mailbox limits." & vbCRLF & "Error Description: " & _
      Err.Description & vbCRLF & "The script will now exit",vbCritical,"Error"
      WScript.Quit
   End If
   
   'Close Objects
   Set objADHelper = Nothing
   Set objOU = Nothing

End Sub