'Created By Matthew Hull on some day a long time ago
'Documented on 5/10/04

'This script will report the size of all the subfolders in a folder.  It will report
'the size in either Gigabytes or Megabytes.  

Option Explicit

On Error Resume Next

Dim objFSO, colFolders, objRootFolder, strPath, objFolder, strOutputFile
Dim txtOutput, bolGBReportOn, intSize, strReportSize

'*************************************************************************************
'Enter the path that you want to report on, and the name and location of the report.
'Then decide if you want the report in Gigabytes (True) or Megabytes (False)

strPath = "F:\"
strOutputFile = "F:\Folder Size.csv"
bolGBReportOn = True

'*************************************************************************************

'Create the File System Object, this will get the folder information
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Verify the folder exists, if not then exit the script
If Not objFSO.FolderExists(strPath) Then
   MsgBox "The folder " & strPath & " doesn't exist.",vbOkOnly,"Folder Doesn't Exist"
   Set objFSO = Nothing
   Wscript.Quit
End If

'Set the Gb vs Mb information
If bolGBReportOn Then
   intSize = 1073741824
   strReportSize = "GB"
Else
   intSize = 1048576
   strReportSize = "MB"
End If

'Create the reportt file
Set txtOutput = objFSO.CreateTextFile(strOutputFile)

'Create the folder object
Set objRootFolder = objFSO.GetFolder(strPath)

'Create a collection of folder objects that contain the subfolders
Set colFolders = objRootFolder.SubFolders

'Write the report by looping thru each folder object in the collection
txtOutput.Write "Folder Name, Size in " & strReportSize & ", Last Modified" & vbCRLF
For Each objFolder in colFolders
   txtOutput.Write objFolder.Name & "," & Round((objFolder.Size/intSize),2) & _
   "," & objFolder.DateLastModified & vbCRLF
Next

'Close the report file
txtOutput.Close

'Display a message to the user when completet
MsgBox "A log of the folder sizes in """ & strPath & """ has been created.  " & _
"You can find it at """ & strOutputFile & """",vbOkOnly,"Log Created"

'Close objects
Set objFSO = Nothing
Set objRootFolder = Nothing
Set colFolders = Nothing