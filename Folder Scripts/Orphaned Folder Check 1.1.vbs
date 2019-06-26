'Created by Matthew Hull on 11/06/03 
'Last Updated 12/29/03 
'Documented on 4/25/04

'Version 1.01

'This script will look for orphaned folders.

'Version History
'~~~~~~~~~~~~~~~
'Version 1.01 - The script now searchs the whole domain for the   

'Version 1.0 - First version of this script released.
'*************************************************************************************

Option Explicit

On Error Resume Next

Dim objFSO, strUser, strRootPath, strPath, intCount, strMessage
Dim objFolder, objRootFolder, colFolders, bolCheck, txtOutput, strOutputFile
Dim objConnection, objCommand, objRootDSE, objRecordSet

'*************************************************************************************
'Set the root folder and the location of the output file

strRootPath = "\\Server\Share"
strOutputFile = "C:\Orphaned Folders.txt"

'*************************************************************************************

'Create a RootDSE object for the domain
Set objRootDSE = GetObject("LDAP://RootDSE")

'Establish a connection to Active Directory using ActiveX Data Object
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADSDSOObject;"

'Create the command object and attach it to the connection object
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

'Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Set the orphan folder count to 0
intCount = 0

'Create a folder object that points to the root
Set objRootFolder = objFSO.GetFolder(strRootPath)

'Create a collection of folders that contain the subfolders of the root
Set colFolders = objRootFolder.Subfolders

'Start the message that will be displayed to the user
strMessage = "The following folders are orphaned in " & strRootPath & vbCRLF

For each objFolder in colFolders

   'objCommand defines the search base (in this case the whole domain) a filter
   '(All object of type user that match the foldername) and some attributes
   'associated with the returned objects (SamAccountName)
   objCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
   ">;(&(objectClass=user)(SamAccountName=" & objFolder.Name & "));SamAccountName"
   
   'Initiate the LDAP query and return results to a RecordSet object.
   Set objRecordSet = objCommand.Execute
   
   'Set the Check to false, it will turn true if a folder isn't orphaned
   bolCheck = False
   
   'Loop thru each item in the Record Set and compare it to the folder name
   'if they match then the folder is not orphaned.  There should only be one
   'item returned in the Record set so this procedure is fast.
   While Not objRecordset.EOF      
      If UCase(objFolder.Name) = UCase(objRecordset.Fields("SamAccountName")) Then
         bolCheck = True
      End If
      objRecordset.MoveNext
   Wend
   objRecordSet.MoveFirst

   'If the folder is orphaned increase the count and add the folder to the message
   If Not bolCheck Then
      intCount = intCount + 1
      strMessage = strMessage & intCount & " - " & objFolder.Name & vbCRLF
   End If
Next

'Close the connection to Active Directory
objConnection.Close

'Change the message if no orphaned folders are located
If intCount = 0 Then
   strMessage = "There are no orphaned folders in " & strRootPath
End If

'Create a text file, write the message to it and close the text file
Set txtOutput = objFSO.CreateTextFile(strOutPutFile)
txtOutput.Write strMessage
txtOutput.Close

'Display a message to the user when complete
MsgBox strMessage,vbOkOnly,"Orpahan Folder Check"

'Close objects
Set objFSO = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRootDSE = Nothing
Set objRecordSet = Nothing