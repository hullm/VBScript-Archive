'Created by Matthew Hull on 9/3/05

'This script will read in a CSV and compair the list of users to users in an OU.  If 
'the user isn't in the CSV then their data is moved to an archive folder and their account
'is deleted.  If you have an OU that has users in it this can be used to remove users that
'aren't needed anymore.  The CSV would contain the users who are supposed to be in the OU
'in a lastname, firstname format
'USE AT YOUR OWN RISK - THIS IS DESIGNED TO DELETE USER ACCOUNTS!!!!

Option Explicit

On Error Resume Next

Dim strSourceCSV, strOU, strArchive, objFSO, txtSourceCSV, objUserList
Dim objADHelper, objOU, objUser

'*************************************************************************************
'Enter the path to the CSV with the list of users that are supposed to be there.
strSourceCSV = "G:\Scripts\Import\Import.csv"

'Enter the OU where the users are supposed to be
strOU = "Students"

'Enter the location where you want the data archived
strArchive = "G:\Shared Data\StudentArchive\04-05"
'*************************************************************************************

'Create the Objects needed to perform the required tasks
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objUserList = CreateObject("Scripting.Dictionary")
Set txtSourceCSV = objFSO.OpenTextFile(strSourceCSV)
Set objADHelper = CreateObject("ADHelper.wsc")

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
   "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Use the ADHelper object to get the OU object
Set objOU = objADHelper.OUObject(strOU)

'See if the archive folder exists, if not exit the script
If Not objFSO.FolderExists(strArchive) Then
   MsgBox "The folder " & strArchive & " doesn't exist.  This script will now exit."
   Wscript.Quit   
End If

'Loop through each line in the CSV and add it to a Dictionary Object.
While txtSourceCSV.AtEndOfLine = False
   objUserList.Add txtSourceCSV.ReadLine, ""
Wend

'Loop thru each user in the OU
For Each objUser in objOU

   'Verify the object is of type user
   If objUser.Class = "user" Then
      If Not objUserList.Exists(objUser.DisplayName) Then
         objFSO.CreateFolder strArchive & "\" & objUser.SamAccountName
         objFSO.CopyFolder objUser.HomeDirectory, strArchive & "\" & objUser.SamAccountName
         objFSO.DeleteFolder objUser.HomeDirectory, True 
         objUser.DeleteObject(0)
      End If
   End If
Next

'Close the CSV file
txtSourceCSV.Close

'Display a message when done
MsgBox "Done"

'Close Open Objects
Set objFSO = Nothing
Set objUserList = Nothing
Set txtSourceCSV = Nothing