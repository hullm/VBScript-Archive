'Created by Matthew Hull on 12/21/04
'Last Updated 01/06/04

'This script will move unwated files out of users home folders.  It will log where the file
'came from in a log file.  You can also have it replace the removed file with a note.

Option Explicit

On Error Resume Next

Dim strFileTypes, strBaseFolder, strDestFolder, strLogFile, bolReplaceFile, strNewFile
Dim objFSO, objBaseFolder, objFile, objFiletypes, intIndex, strTempFileType, bolStartLog
Dim strFileType, strFileName, txtOutput, objAccessDenied, strDeniedFolder, strLog
Dim bolLogOnly, intSizeofRemovedData, datStartTime, datEndTime, intStartTimer
Dim intEndTimer, intTotalTime, intSeconds, intFolderCount, intTotalFileCount
Dim intUnwantedFileCount, intTotalSizeofData, intSizePercent, intFilePercent
Dim intDeniedFolderCount, strMessage, intDeniedFileCount, objSourceFolders

Set objSourceFolders = CreateObject("Scripting.Dictionary")

'****************************************************************************************
'All user entry is done in this section

'If you want to just log the unwanted files instead of moving them set this to true
bolLogOnly = True

'Enter the file extensions you want to scan for separated with a comma
strFileTypes = "mp3,wma,m3u,pls,asx,mpeg,iso,exe,avi,mov,ogg,aac,wav,msi,vbs,bat,mpg,smc,nes,smd,d64,bin,js"

'Enter the complete path to the folder you want to scan
objSourceFolders.Add "D:\Shared Date\2005",""
objSourceFolders.Add "D:\Shared Date\2006",""
objSourceFolders.Add "D:\Shared Date\2007",""
objSourceFolders.Add "D:\Shared Date\2008", ""

'Enter the complete path to the location you want to move all the unwanted files
strDestFolder = "D:\Reports\Unwanted Files"

'Enter the path to the log file.  It will add to it each time the script is run
strLogFile = "D:\Reports\Unwanted Files.log"

'If you want to replace the file with a note set this to true and choose a file.
bolReplaceFile = False
   strNewFile = "D:\Reports\ReadMe.Doc"
   
'****************************************************************************************
'DO NOT CHANGE ANYTHING BELOW THIS LINE 
'****************************************************************************************

For Each strBaseFolder in objSourceFolders
   Main(strBaseFolder)
Next   

MsgBox "Complete"

'****************************************************************************************

Sub Main(strBaseFolder)

   'This is the main Subroutine 

   On Error Resume Next

   'Create the File System Object                         
   Set objFSO = CreateObject("Scripting.FileSystemObject")

   VerifyFolders
   CreateLogFile  
   BuildFileTypeDO   

   'If we are only logging we don't want to copy a replacement file over.
   If bolLogOnly Then
      bolReplaceFile = False
   End If
      
   'Create the Dictionary Object that will hold a list of folders that we are denied access too
   Set objAccessDenied = CreateObject("Scripting.Dictionary")
   
   'Create the Folder Object for the base folder
   Set objBaseFolder = objFSO.GetFolder(strBaseFolder)
   
   'initialize variables
   intTotalSizeofData = 0
   intSizeofRemovedData = 0
   intFolderCount = 0
   intTotalFileCount = 0
   intUnwantedFileCount = 0
   intDeniedFileCount = 0
   datStartTime = Time
   intStartTimer = Timer
   
   'Move all the unwanted files in the root of the folder
   MoveFiles(objBaseFolder)
   
   'Scan for unwanted files recursively
   SubFolderScan(objBaseFolder)
   
   'Comlete Time
   datEndTime = Time
   intEndTimer = Timer
   
   CalculateStats
   
   'Generate message to display to the user and place in the log.
   strMessage = intSizeofRemovedData & " MB's out of " & intTotalSizeofData & " MB's were unwanted.  " & _
      "This represents " & intSizePercent & "% of the data" & vbCRLF & _
      intUnwantedFileCount & " files out of " & intTotalFileCount & " were unwanted.  " & _
      "This represents " & intFilePercent & "% of the files" & vbCRLF & _
      "You were denied access to " & intDeniedFolderCount & " folders out of " & intFolderCount & "." & vbCRLF & _
      "You were denied access to " & intDeniedFileCount & " files out of " & intTotalFileCount & "." & vbCRLF & _
      "Completed in " & intTotalTime & vbCRLF
      
   CloseLogFile
      
   'Close Open Objects
   Set objFSO = Nothing
   Set objFileTypes = Nothing
   Set objBaseFolder = Nothing
   
End Sub

'****************************************************************************************

Sub VerifyFolders

   'Verify the Base and Dest folders exist
   
   If Not objFSO.FolderExists(strBaseFolder) Then
      MsgBox "The Base Folder doesn't exisit" & vbCRLF & strBaseFolder,vbCritical,"Missing Folder"
      WScript.Quit
   End If
   If Not objFSO.FolderExists(strDestFolder) Then
      MsgBox "The Destination Folder doesn't exisit" & vbCRLF & strDestFolder,vbCritical,"Missing Folder"
      WScript.Quit
   End If
   
   'Add a \ to the end of strDestFolder if it doesn't have one.
   If Right(strDestFolder,1) <> "\\" Then
      strDestFolder = strDestFolder & "\"
   End If

End Sub

'****************************************************************************************

Sub CreateLogFile

   'Open or create the log file.  The bolStartLog is set to False to start with.  If an 
   'unwanted file is found that will be changed.  That will trigger strLog to be written 
   'to the log file.  This way if no unwanted files are found nothing will be written to
   'the log file.  If this script is scheduled to run nightly it will make for a cleaner
   'log.
   
   Set txtOutput = objFSO.OpenTextFile(strLogFile,8,True)
   If Err Then
      MsgBox "Could Not Create the Log File.",vbCritical,"Error Detected"
      Err.Clear
      WScript.Quit
   End If
   strLog = "Scan started on " & Date & " at " & Time & vbCRLF
   bolStartLog = False
   

End Sub

'****************************************************************************************

Sub BuildFileTypeDO

   'Add each file type to a Dictionary object
   Set objFileTypes = CreateObject("Scripting.Dictionary")
   
   If objFileTypes.Count = 0 Then
      For intIndex = 1 to Len(strFileTypes)
         If Mid((strFileTypes),intIndex,1) <> "," Then
            strTempFileType = strTempFileType & Mid((strFileTypes),intIndex,1)
         Else
            objFiletypes.Add UCase(strTempFileType),""
            strTempFileType = ""
         End If   
      Next
      objFiletypes.Add UCase(strTempFileType),""
   End If

End Sub

'****************************************************************************************

Sub SubFolderScan(objFolder)

   'This will do a recursive folder/file scan. 
   
   On Error Resume Next
   
   Dim objSubFolder, colFolders
   
   intFolderCount = intFolderCount + 1
         
   'Create a Collection object that contains all the subfolders
   Set colFolders = objFolder.SubFolders
   
   'Move all the unwanted files in the subfolders
   For Each objSubFolder in colFolders
      If Err Then
         intFolderCount = intFolderCount - 1
         Err.Clear
         Exit For
      End If
      Call MoveFiles(objSubFolder)
      Call SubFolderScan(objSubFolder)
   Next                                         
   
End Sub

'****************************************************************************************

Sub MoveFiles(objSubFolder)

   On Error Resume Next
   
   Dim colFiles
   
   'Create a collection object that contains all the files in the folder
   Set colFiles = objSubFolder.Files
   
   'Loop through each file, if it is unwanted then move it
   For Each objFile in colFiles
   
      'Their would be an error if access was denied to the folder, if so add the folder
      'to the AccessDenied Dicionary object and exit the for loop.
      If Err Then
         objAccessDenied.Add objSubFolder.Path, ""
         Err.Clear         
         Exit For
      End If
      
      'Increase the file count by 1 and add the size of the file to the total size
      intTotalFileCount = intTotalFileCount + 1
      intTotalSizeofData = intTotalSizeofData + objFile.Size
      
      'Loop through each file type in the FileType Dictionary object and compair it to 
      'the last three letters in the file.  If a match is found the move the file.
      For Each strFileType in objFileTypes 
         If UCase((Right(objFile.Path,Len(strFileType)))) = strFileType Then

            'The file is unwated so increase the counter by one and add the size of the
            'file to the SizeofRemovedData variable.  Then get the file name.
            intUnwantedFileCount = intUnwantedFileCount + 1
            intSizeofRemovedData = intSizeofRemovedData + objFile.Size
            strFileName = objFile.Path
            
            'If we are not running this script in log only mode then it will attempt to
            'move the data.  If it fails on the move it will copy then force a delete.
            If Not bolLogOnly Then
               objFile.Move(strDestFolder)
               If Err Then
                  Err.Clear
                  objFile.Copy(strDestFolder)
                  objFile.Delete True                  
               End If
            End If
            
            'bolStartLog is turned off by default.  Since an unwanted file is found we
            'want to change this to true so data will be written to the log.  The header
            'for the logfile will be written to it.
            If Not bolStartLog Then
               txtOutput.Write strLog
               bolStartLog = True
            End If
            
            'Write the name of the file to the log, first look for errors, if there is one then
            'it couldn't move or copy the folder.  It will see if the delete was successful and
            'report the status to the log.
            If Err Then
               If objFSO.FileExists(strFileName) Then
                  txtOutput.Write "Access was denied, file NOT removed: " & strFileName & vbCRLF
               Else
                  txtOutput.Write "Access was denied, file removed: " & strFileName & vbCRLF
               End If
               intDeniedFileCount = intDeniedFileCount + 1
               Err.Clear
            Else
               txtOutput.Write strFileName & vbCRLF
            End If
               
            'If we want to replace the file then copy the source file over with the old
            'files name, keeping the source files extension.
            If bolReplaceFile Then
               strFileName = (Left(strFileName,(Len(strFileName)-3))) & Right(strNewFile,3)
               Call objFSO.CopyFile(strNewFile, strFileName)
            End If
            
            'Once the file is recognized as an unwanted file then exit the for loop, no
            'need to check the file that has already been removed against other file types
            Exit For
         End If
      Next
   Next 
End Sub

'****************************************************************************************

Sub CalculateStats

   'Calculate the length of time it took to run the script
   intTotalTime = intEndTimer - intStartTimer
   intSeconds = intTotalTime Mod 60
   intTotalTime = Int(intTotalTime/60)
   If Len(intSeconds) <> 2 Then
      intSeconds = "0" & intSeconds
   End If
   intTotalTime = intTotalTime & "m " & intSeconds & "s"
   
   'Calculate percents
   intFilePercent = Round(((intUnwantedFileCount/intTotalFileCount)*100),2)
   intSizePercent = Round(((intSizeofRemovedData/intTotalSizeofData)*100),2)
   
   'Calculate the size of the data in MB
   intSizeofRemovedData = Round((intSizeofRemovedData/1048576),2)
   intTotalSizeofData = Round((intTotalSizeofData/1048576),2)
   
   'Count the number of denied folders
   intDeniedFolderCount = objAccessDenied.Count
   intFolderCount = intFolderCount + intDeniedFolderCount

End Sub

'****************************************************************************************

Sub CloseLogFile
   
   'Close the logfile
   For Each strDeniedFolder in objAccessDenied
      If Not bolStartLog Then
         txtOutput.Write strLog
         bolStartLog = True
      End If       
      txtOutput.Write "*** Access Denied *** " & strDeniedFolder & vbCRLF
   Next
   If bolStartLog Then    
      txtOutput.Write strMessage
      txtOutput.Write "***********************************************************************" & vbCRLF
   End If
   txtOutput.Close

End Sub