'Created by Matthew Hull on 11/18/03
'Last Updated 12/31/03

'Version 1.02

'This script will scan an OU and generate a report in a CSV format.

'Version History
'~~~~~~~~~~~~~~~
'Version 1.02 - Now you just need to enter the name of the OU, not the LDAP
'               path.  A prompt will be displayed if more then one OU with
'               the nane exists and you pick the one you want.
'             - Optimized orphan folder check, it's now a lot faster.
'             - The users home folder location will be pulled from the user
'               accounts automatically.  If more then one is located you will
'               be asked which one you want.
'             - Fixed bug with bolHFPathCheck, removed bolFolderCheck
'             - Code has been documented.
'             - Username format is now discovered by the script.
'             - The script will now determine what reports to turn on by scanning 
'               to see if each property has a value.  If a user does that report 
'               is turned on.
'             - A bug with strFolder is fixed.  Previously if the user had placed 
'               a "\" at end of strFolder the script would have failed, now it
'               looks for it and removes if found.           
'             - Inputs have been added so you no longer have to edit the script
'               you can just run it and it will ask you for what it needs.
'             - strTitle defaults to the OU name and strFolder defualts to the 
'             - Desktop folder

'Version 1.01 - Added the ability to choose weather you want to scan the OU or
'               the whole domain when doing an orphan check.  
'             - Changed the text to Access Denied from a null character in the 
'               directory check section.
'             - Moved all the data entry to the top of the script.
'             - Fixed a bug with the bolFolderCheck variable.

'Version 1.0 - First version of this script released.

'This is were all variables are declaired.
Option Explicit
Dim objOU, objFSO, txtOutput, strOutputFile, objUser, strOU, strTitle, strFolder
Dim strHomeFolderRoot, objHomeFolder, strFullOU, RegExp, colFolders, objFolder
Dim bolCheck, objRootDSE, bolLastNameFirst, strUserName, strDirSize, bolName, strDN
Dim bolUserName, bolHomeFolder, bolProfile, bolScript, bolDescription, bolSearchDomain
Dim bolDirSize, bolUserNameCheck, bolOrphanCheck, bolHFPathCheck, strDerivedHomeFolder
Dim objConnection, objCommand, objRecordSet, strInput, intCount, objOUList, intChoice
Dim strFullOUNoCommas, objHomeFolders, Key, intLengthofHomeFolder, bolKeyExist
Dim strLastNameFirst, strFirstNameFirst, intLastNameFirst, intFirstNameFirst
Dim objShell, colSpecialFolders, intTotalNumberofUsers
Set objFSO = CreateObject("scripting.FileSystemObject")

'Enter the OU name.
Do 
   strOU = InputBox  ("What OU would you like to scan?","OU Name"," ")
   If strOU = "" Then
      WScript.Quit
   ElseIf strOU = " " Then
      MsgBox "You must specify an OU",vbOkOnly,"Name Required"
   End If
Loop Until strOU <> " "

'Enter the title of the report, this will also be the file name. OU name is Default.
strTitle = InputBox  ("What would you like this report called?","Report Name",strOU)
If strTitle = "" Then
   WScript.Quit
End If

'Choose where the reports will be stored and verify location exists.  Desktop is default.
Do Until objFSO.FolderExists(strFolder)
   Set objShell = CreateObject("WScript.Shell")
   Set colSpecialFolders = objShell.SpecialFolders
   strFolder = colSpecialFolders.Item("Desktop")
   strFolder = InputBox("Where would you like this report saved?", _
   "Report Location",strFolder)
   If strFolder = "" Then
      WScript.Quit
   ElseIf Not objFSO.FolderExists(strFolder) Then
      MsgBox "Output folder " & strFolder & _
      " does not exist.",vbOkOnly,"Folder Error"
   End If
Loop

'Call the Report Subroutine 
Report

'This message is displayed when the report is done.
MsgBox "The report is complete and can be found in " & _
strFolder & ".",vbOkOnly,"Report Complete"

Sub Report   
   On Error Resume Next
   
   'This Regular Expression object will be used thru out the script to 
   'replace characters in strings.
   Set RegExp = New RegExp
   RegExp.Global = True
   RegExp.IgnoreCase = True

   'This section of code is used to retrieve a list of OU's from the domain
   'that match the requested name.  If there is only one returned the script
   'will continue to run.  If more then one OU is returned it will display a
   'list of OU's for you to choose from.

   'Establish a connection to Active Directory using ActiveX Data Object
   Set objConnection = CreateObject("ADODB.Connection")
   objConnection.Open "Provider=ADSDSOObject;"

   'Create the command object and attach it to the connection object
   Set objCommand = CreateObject("ADODB.Command")
   objCommand.ActiveConnection = objConnection

   'objRootDSE is created to determine the domain name.  If it can't connect
   'to the domain the script will exit. objCommand defines the search base 
   '(in this case the whole domain) a filter (All object of type 
   'OrganizationalUnit) and some attributes associated with the returned 
   'objects (Name and DistinguishedName)
   Err.Clear
   Set objRootDSE = GetObject("LDAP://rootDSE")
   If Err Then
      MsgBox "There was an error connecting to your domain.  Make sure you are " & _
      "connected to the network, or you might have to run this from a Domain " & _
      "Controller",vbOkOnly,"Cannot Connect to Domain"
      WScript.Quit
   End If   
   objCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
   ">;(&(objectClass=OrganizationalUnit));Name,DistinguishedName"

   'Initiate the LDAP query and return results to a RecordSet object.
   Set objRecordSet = objCommand.Execute

   'Create the object objOUList which will contain all the matching OU's
   Set objOUList = CreateObject("Scripting.Dictionary")
   
   'intCount will contain the number of OU's returned that match the requested OU.
   'intChoice will be 1 if only one OU is found and chossen by the user if more 
   'then one OU is found.  This will be used to select the OU from objOUList.
   intCount = 0
   intChoice = 0

   'If more then one OU is found this will be the error message displayed.
   strInput = "There was more then one OU found with named " & strOU & ", please "
   strInput = strInput & "choose the correct OU from the list below" & vbCRLF
   strInput = strInput & "The input will be required on the next screen."
   strInput = strInput & vbCRLF & vbCRLF

   'Loop thru the returned RecordSet and look for the OU.  If found it adds the 
   'DN name to objOUList and the error message. It will also increases intCount by 1.
   While Not objRecordset.EOF
      If uCase(objRecordset.Fields("Name")) = UCase(strOU) Then
         intCount = intCount + 1
         strInput = strInput & intCount & " - " & _
         objRecordset.Fields("DistinguishedName") & vbCRLF
         strDN = objRecordSet.Fields("DistinguishedName")
         objOUList.Add "Key" & intCount, strDN
      End If
      objRecordset.MoveNext
   Wend

   'Close the connection with Active Directory
   objConnection.Close

   'If more then one OU is found it will ask you which you want, if only one is
   'found it will set intChoice to 1.  If no OU is found the script will exit.
   If intCount > 1 Then
      Do Until objOUList.Item("Key" & intChoice) <> ""
         MsgBox strInput,vbOkOnly,strTitle
         intChoice = InputBox("Please choose your OU","Select OU","Enter " & _
         "a number from the previous screen")
         If intChoice = "" Then
            WScript.Quit
         ElseIf objOUList.Item("Key" & intChoice) = "" Then
            MsgBox "You must choose a valid OU to continue.",vbOkOnly,"Error"
         End If
      Loop 
   ElseIf intCount = 0 Then
      MsgBox strOU & " is not a valid OU.",vbOkOnly,"Invalid OU"
      WScript.Quit
   Else
      intChoice = 1
   End If

   'Build the OU string and create the OU object
   strFullOU = "LDAP://" & objOUList.Item("Key" & intChoice)
   Set objOU = GetObject(strFullOU)

   'Uses a regular expression to assist with the removal of commas in strFullOUNoCommas
   'The commas are removed for the title of the report.  A comma separates each field so
   'they are removed so it will not split over multiple cells.
   RegExp.Pattern = ","
   strFullOUNoCommas = RegExp.Replace(strFullOU," ")   

   'Now that we are done selecting the OU we will start to look for a valid home folder
   'within each user in the OU.  This script assumes that the home folder in the
   'account will end with the username of the user.  It will take the home folder and
   'subtract the number of characters in the SamAccountName + 1 for the "\".  The
   'remaining path will be added to a dictionary object if it is found to be unique.
   'When complete if more then one folder is found in the dictionary object it will ask
   'you which one you want.  If only one is found it will continue on.

   'This dictionary object will store the home folders found in the OU.
   Set objHomeFolders = CreateObject("Scripting.Dictionary")

   'If more then one home folder is found this will be the error message displayed.
   strInput = "There was more then one home folder found in " & strOU & ", please "
   strInput = strInput & "choose the correct home folder from the list below" & vbCRLF
   strInput = strInput & "The input will be required on the next screen."
   strInput = strInput & vbCRLF & vbCRLF
   objHomeFolders.Add "Key0", "No Folder - Turn off all folder checks"
   strInput = strInput & "0 - No Home Folders" & vbCRLF

   'intCount will contain the number of home folders returned, intChoice will be 1 if 
   'only one home folder is found and chosen by the user if more then one home folder is
   'found.  This will be used to select the home folder from objHomeFolders.
   intCount = 0
   intChoice = 0

   'Loop thru the OU and add each unique home folder to the objHomeFolders, then increase 
   'inCOunt by 1 and build the error message incase more then one home folder is found.
   For Each objUser in objOU
      If objUser.Class = "user" Then
         intLengthofHomeFolder = (Len(objUser.HomeDirectory)) - (Len(objUser.SamAccountName))
         intLengthofHomeFolder = Abs(intLengthofHomeFolder) - 1
         strDerivedHomeFolder = Left(objUser.HomeDirectory,intLengthofHomeFolder)
         bolKeyExist = False
         For Each Key in objHomeFolders
            If LCase(objHomeFolders.Item(Key)) = LCase(strDerivedHomeFolder) Then
               bolKeyExist = True
            End If
         Next
         If Not bolKeyExist Then
            If strDerivedHomeFolder <> "" Then
               intCount = intCount + 1
               objHomeFolders.Add "Key" & intCount, strDerivedHomeFolder
               strInput = strInput & intCount & " - " & strDerivedHomeFolder & vbCRLF
            End If
         End If
      End IF
   Next

   'If more then one home folder is found it will ask you which you want, if only one is
   'found it will set intChoice to 1.
   If intCount > 1 Then
      Do 
         MsgBox strInput,vbOkOnly,strTitle
         intChoice = InputBox("Please choose your home folder","Select Home Folder","Enter " & _
         "a number from the previous screen")
         If intChoice = "" Then
            WScript.Quit
         ElseIf objHomeFolders.Item("Key" & intChoice) = "" Then
            MsgBox "You must choose a valid home folder to continue.",vbOkOnly,"Error"
         End If
      Loop Until objHomeFolders.Item("Key" & intChoice) <> ""
   ElseIf IntCount = 0 Then
      intChoice = 0
   Else
      intChoice = 1
   End If

   'If no home folders are found this will turn off all checks related to home folders, otherwise
   'it will set strHomeFolders to the chosen home folder and turn on all checks.
   If (objHomeFolders.Count = 0) or (intChoice = 0) Then
      bolHomeFolder = False
      bolHFPathCheck = False
      bolDirSize = False
      bolOrphanCheck = False
   Else
      strHomeFolderRoot = objHomeFolders.Item("Key" & intChoice)
      bolHomeFolder = True
      bolHFPathCheck = True
      bolDirSize = True
      bolOrphanCheck = True
      bolSearchDomain = True 'The only place where this can be changed.
   End If
   
   'Now we are done gathering home folder location information we will loop thru the OU and
   'decide what we want in our report.  If a value for a certain component exists it will be added
   'to the report.  Before it starts it will set all the components to False.  If found in the OU
   'the value will be changed to True.

   bolName = False
   bolUsername = False
   bolProfile = False
   bolScript = False
   bolDescription = False
   intLastNameFirst = 0
   intFirstNameFirst = 0
   intTotalNumberofUsers = 0
    
   For Each objUser in objOU
     If objUser.Class = "user" Then
        intTotalNumberofUsers = intTotalNumberofUsers + 1
        If (objUser.sn <> "") Or (objUser.GivenName <> "") Then
           bolName = True
        End If
        If objUser.SamAccountName <> "" Then
           bolUserName = True
        End If
        If objUser.ProfilePath <> "" Then
           bolProfile = True
        End If
        If objUser.ScriptPath <> "" Then
           bolScript = True
        End If
        If objUser.Description <> "" Then
           bolDescription = True
        End If
        If bolUserName And bolName Then
           'This will try to build the username in both formats and check it against the real
           'username, the one that is correct will be increased by one.  The winner becomes the
           'username format.
           bolUserNameCheck = True
           strLastNameFirst = LCase(Trim(objUser.sn)) & LCase(Left(Trim(objUser.GivenName),1))
           strFirstNameFirst = LCase(Left(Trim(objUser.GivenName),1)) & LCase(Trim(objUser.sn))
           If strLastNameFirst = LCase(objUser.SamAccountName) Then
              intLastnameFirst = intLastNameFirst + 1
           ElseIf strFirstNameFirst = LCase(objUser.SamAccountName) Then
              intFirstNameFirst = intFirstNameFirst + 1
           End If
        End If
     End If
   Next
   
   'This is where we check to see which format was the dominant format.  If neither are correct 
   'the username check will be turned off.  It requires that there be at least 50% of the
   'usersnames creatable from the OU.
   If bolUserNameCheck Then
      If ((intLastNameFirst/intTotalNumberofUsers)*100 > 50) Then
         bolLastNameFirst = True
      ElseIf ((intFirstNameFirst/intTotalNumberofUsers)*100 > 50) Then
         bolLastNameFirst = False
      Else
         bolUserNameCheck = False
      End If
   End If
            
   'Now that we know what will be in the report we can start to write it.
   
   'Create the output file, if there is an error the script will exit.  The regular
   'expression will assist with removing "\\" from the folder location.
   Err.Clear
   RegExp.Pattern = "\\\\"
   strOutputFile = strFolder & "\" & strTitle & ".csv"
   strOutputFile = RegExp.Replace(strOutputFile,"\")
   'This next If statment will add the "\" back to the beginning of a UNC path.
   If Left(strOutPutFile,1) = "\" Then
      strOutPutFile = "\" & strOutPutFile
   End If
   Set txtOutPut = objFSO.CreateTextFile(strOutputFile)
   If Err Then
      MsgBox "There was an error creating " & strOutputFile & vbCRLF & "Description: " & _
      Err.Description & vbCRLF & "Error Number: " & Err.Number,vbOkOnly,"Cannot Create File"
      WScript.Quit
   End If

   'This section will create the title portion of the report.  If the value of the boolean 
   'variable is set to true the title for that compnent will be added.      
   txtOutput.Write strTitle & ",,,Home Folder Location = " & strHomeFolderRoot & vbCRLF & _
   ",,,OU Location: " & strFullOUNoCommas & vbCRLF
   If bolName Then
      txtOutput.Write "Last Name,First Name"
   End If
   If bolUserName Then
      txtOutput.Write ",Username"
   End If
   If bolHomeFolder Then 'This section has two comma's to accommodate the drive letter
      txtOutput.Write ",,Home Folder" 
   End If
   If bolProfile Then
      txtOutput.Write ",Profile"
   End If
   If bolScript Then
      txtOutput.Write ",Script"
   End If
   If bolDescription Then
      txtOutput.Write ",Description"
   End If
   If bolDirSize Then
      txtOutput.Write ",Dir Size MB"
   End If
   If bolUserNameCheck Then
      txtOutput.Write ",UN Check"
   End If
   If bolHFPathCheck Then
      txtOutput.Write ",HF Path Check"
   End If
   txtOutput.Write vbCRLF

   'The next section will add the data to the report.

   For Each objUser in objOU
      If objUser.Class = "user" Then          
         If bolName Then 'Writes last name, first name
            txtOutput.Write objUser.sn & "," & objUser.GivenName 
         End If         
         If bolUserName Then 'Writes username
            txtOutput.Write "," & objUser.SamAccountName 
         End If
         If bolHomeFolder Then 'Writes home drive letter, and home folder location
            txtOutput.Write "," & objUser.HomeDrive & "," & objUser.HomeDirectory
         End If
         If bolProfile Then 'Writes profile location
            txtOutput.Write "," & objUser.ProfilePath
         End If
         If bolScript Then 'Writes script
            txtOutput.Write "," & objUser.ScriptPath
         End If
         If bolDescription Then 'Writes descriptions and removes commas.
            RegExp.Pattern = ","
            txtOutput.Write "," & RegExp.Replace(objUser.Description, "") 
         End If
         If bolDirSize Then
            'The Err object is used to see if access is denied on the folder.  If you don't have 
            'access to the folder an error will be generated.  This line will clear the error
            'before the check just in case there was another error before.
            Err.Clear 

            'Check and see if a folder exists for the user.  If no folder exists "No Folder" will
            'be written to the report.  If a folder is found it will attempt to get the size, and
            'write it to the report.  If an error is generated then "Access Denied" is written to
            'the report and the Err object is cleared.
            If Not objFSO.FolderExists(strHomeFolderRoot & "\" & objUser.SamAccountName) Then
               strDirSize = "No Folder"
            Else
               Set objHomeFolder = objFSO.GetFolder(strHomeFolderRoot & _
               "\" & objUser.SamAccountName)
               strDirSize = Round((objHomeFolder.Size/1048576),2)
            End If             
            If Err Then
               strDirSize = "Access Denied"
               Err.Clear
            End If
            txtOutput.Write "," & strDirSize
         End If
         If bolUserNameCheck Then
            'The username check will look for two different username formats.  First letter of first
            'name then last name or last name then first letter of first name. It will pull this
            'information from the first and last name values stored in Active Directory.  If the
            'username is found to be invalid it will write "Bad UserName" to the report.
            If bolLastNameFirst Then
               strUserName = LCase(Trim(objUser.sn)) & LCase(Left(Trim(objUser.GivenName),1))
            Else
               strUserName = LCase(Left(Trim(objUser.GivenName),1)) & LCase(Trim(objUser.sn))
            End If
            If Lcase(Trim(objUser.SamAccountName)) <> strUserName Then
               txtOutput.Write ",Bad UserName"
            Else
               txtOutput.Write ","
            End If      
         End If
         If bolHFPathCheck Then
            'If the user has a value for home directory location stored in Active Directory it will
            'verify that it matches the home folder for other users in the OU.  If not it will write
            'Bad Home Folder" to the report
            If objUser.HomeDirectory <> "" Then
               If LCase(objUser.HomeDirectory) <> LCase(strHomeFolderRoot & _
               "\" & objUser.SamAccountName) Then
                  txtOutput.Write ",Bad Home Folder"
               Else
                  txtOutput.Write ","
               End If   
            Else
               txtOutput.Write ","
            End If
         End If
         txtOutput.Write vbCRLF
      End If
   Next

   'The main part of the report is done at this point, now we just have the orphan check.
   'The orphan check will establish a connection to Active Directory then look up each
   'folder name in AD.  If it isn't found it will consider the folder orphaned.  There are
   'two modes you can run the check in.  The first will search the whole domain for a match,
   'the second will scan just the OU.

   If bolOrphanCheck Then
      'Write the Orphaned Folder title to the report
      txtOutput.Write vbCRLF & "Orphaned Folders" & vbCRLF

      'Create a folder object then a collection object that contains all the folders to scan
      Set objHomeFolder = objFSO.GetFolder(strHomeFolderRoot)
      Set colFolders = objHomeFolder.SubFolders

      'Establish a connection to Active Directory using ActiveX Data Object
      Set objConnection = CreateObject("ADODB.Connection")
      objConnection.Open "Provider=ADSDSOObject;"

      'Create the command object and attach it to the connection object
      Set objCommand = CreateObject("ADODB.Command")
      objCommand.ActiveConnection = objConnection       

      For Each objFolder in colFolders
         If bolSearchDomain Then

            'objCommand defines the search base (in this case the whole domain) a filter
            '(All object of type user that match the foldername) and some attributes
            'associated with the returned objects (SamAccountName)
            objCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
            ">;(&(objectClass=user)(SamAccountName=" & objFolder.Name & "));SamAccountName"
         
         Else

            'objCommand defines the search base (in this case the OU) a filter
            '(All object of type user that match the foldername) and some attributes
            'associated with the returned objects (SamAccountName)
            objCommand.CommandText = "<" & strFullOU & ">;(&(objectClass=user)" & _
            "(SamAccountName=" & objFolder.Name & "));SamAccountName"

         End If 

         'Initiate the LDAP query and return results to a RecordSet object.
         Set objRecordSet = objCommand.Execute

         'Check to see if the returned user matches the folder name, if not it is marked
         'as orphaned
         bolCheck = False
         While Not objRecordset.EOF      
            If UCase(objFolder.Name) = UCase(objRecordset.Fields("SamAccountName")) Then
               bolCheck = True
            End If
            objRecordset.MoveNext
         Wend
         If Not bolCheck Then
            txtOutput.Write objFolder.Name & vbCRLF
         End If
      Next

      'Close the connection with Active Directory
      objConnection.Close
   End If

   'Close the report file
   txtOutput.Close

   'Close all open objects
   Set txtOutput = Nothing
   Set objOU = Nothing
   Set RegExp = Nothing
   Set objConnection = Nothing
   Set objCommand = Nothing
   Set objRootDSE = Nothing
   Set objRecordSet = Nothing
   Set objShell = Nothing
   Set colSpecialFolders = Nothing
End Sub

Set objFSO = Nothing