'Created by Matthew Hull 5/20/04
'Last modified 9/9/05

'Version 1.01

'This script will reset the passwords for all users in a specified OU.
'You can have it assign the same password to all users or generate a 
'random password for each user.

'Version History
'~~~~~~~~~~~~~~~
'Version 1.01 - Improved error detection.

'Version 1.0 - First version of this script released.

Option Explicit

On Error Resume Next

Dim intPasswordLength, strOU, objUser, objOU, objADHelper, intRandomNumber, txtLogFile
Dim strPassword, bolForceChange, strLogFile, objFSO, bolNeverExpire

'*****************************************************************************************
strOU = "Script Test" 'Name of the OU that contains the users you want to modify
strLogFile = "C:\Password Change Log.csv" 'Location and name of the log
bolForceChange = False 'Set True to have the user to change their password on next logon
bolNeverExpire  = True 'Set True to set the password never to expire
intPasswordLength = 8 'Set the length of the random password
strPassword = "" 'Set generic password for all users if inPasswordLength = 0
'*****************************************************************************************

'Exit the script if bolForceChange and bolNeverExpire are set to true.
If bolForceChange And bolNeverExpire Then
   MsgBox "bolForceChange and bolNeverExpire cannot both be True." & vbCRLF & _
   "The script will now exit.",vbCritical,"Error"
   WScript.Quit
End If

'Create the OU Helper object and get the OU object using the OU Helper object   
Set objADHelper = CreateObject("ADHelper.wsc")

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
      "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Use the ADHelper object to create the OU object
Set objOU = objADHelper.OUObject(strOU)

'Exit if the there is an error creating the OU object
If Err Then
   MsgBox "There is a problem with your strOU setting, OU Not Found" & _
      "The script will now exit.",vbCritical,"OU Not Found"
   Err.Clear
   WScript.Quit
End If

'Create the File System Object, and use it to create the log file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = objFSO.CreateTextFile(strLogFile)
txtLogFile.WriteLine("Display Name,User Name,Password")
If Err Then ErrorMessage

'Loop through each user and set the password
For Each objUser in objOU
   If intPasswordLength <> 0 Then
      strPassword = GeneratePassword(intPasswordLength)
   End If
   If objUser.class = "user" Then
      Call SetPassword(objUser, strPassword, bolForceChange, bolNeverExpire)
      txtLogFile.WriteLine("""" & objUser.DisplayName & """," & _
         objUser.SamAccountName & "," & strPassword)
   End If   
Next

MsgBox "The passwords have been reset.",vbOkOnly,"Done"

'Close the log file
txtLogFile.Close
'*****************************************************************************************
Sub SetPassword(objUser,strPassword,bolForceChange,bolNeverExpire)
   Dim intUAC
   objUser.SetPassword(strPassword)
   objUser.SetInfo
   objUser.Put "UserAccountControl", 512 ' Normal Account
   objUser.SetInfo
   If bolForceChange Then
      objUser.Put "pwdLastSet", 0 'User must change password at next login          
      objUser.SetInfo        
   End If
   intUAC = 512
   If bolNeverExpire Then
      intUAC = intUAC + 65536
   End If
   objUser.Put "UserAccountControl", intUAC
   objUser.SetInfo
End Sub
'*****************************************************************************************
Function GeneratePassword(intLength)
   Dim intIndex, intRandomNumber
   Randomize Timer
   For intIndex = 1 to intLength
      intRandomNumber = Int(2 * Rnd)   
      Select Case intRandomNumber
         Case 0 
            intRandomNumber = Int(10 * Rnd) + 48 '0-9
         Case 1
            intRandomNumber = Int(26 * Rnd) + 97 'a-z
      End Select   
      GeneratePassword = GeneratePassword & Chr(intRandomNumber)
   Next
End Function
'*****************************************************************************************
Sub ErrorMessage
   MsgBox "There was an error detected." & vbCRLF & vbCRLF & "Error Description: " & _
      Err.Description & vbCRLF & "Error Number: " & Err.Number & vbCRLF & vbCRLF & _
      "The script will now exit...",vbCritical,"Error"
   WScript.Quit
End Sub
'*****************************************************************************************