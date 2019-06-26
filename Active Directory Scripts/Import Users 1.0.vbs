'Created by Matthew Hull 4/29/04
'Last modified 8/15/04

'This script will import users into Active Directory from a CSV file, then
'create home folders, setup permissions and add the users to groups.
'The format of the CSV should be LastName,FirstName on each line.  The script
'will verify the CSV format before it runs.  If it finds an error it will let
'you know which line is incorrect.

Option Explicit

On Error Resume Next

Dim bolReportOnly, strSourceCSV, strLogFile, bolLastNameFirst, strOU 
Dim strHomeFolderRoot, strHomeDrive, strProfilePath, strScript, strDescription
Dim strGroups, intPasswordLength, strPassword, strFolderPermissions, strInitial 
Dim strOwnerPermission, objVerifiedGroups, objPerms, objOU, intNewUserCount
Dim intNewFolderCount, txtSourceCSV, strImportedData, intIndex, objNewUser
Dim strFirstName, txtLogFile, strOutput, strLastName, strDefaultNamingContext 
Dim strDomain, strNameSuffix, strNetBIOSDomain, objGroup, objGroups, strError

'*****************************************************************************************
'*                 All user input is done in the following section                       *
'*****************************************************************************************

'To generate a report only and not import users set bolReportOnly to True.  If you want
'the users to be imported set it to False
bolReportOnly = False

'Required Entries
'~~~~~~~~~~~~~~~~
'Location of the import CSV file, file must be in lastname,firstname format
strSourceCSV = "C:\import\import.csv"

'Location where you want the log created.  This folder path must exist prior to running
'the script
strLogFile = "C:\import\New User Log.csv"

'Username format, If the users name is John Smith and you want the username to be
'smithj then set bolLastNameFirst to True.  If you want the username to be jsmith then
'set bolLastNameFirst to False
bolLastNameFirst = False

'Enter the name of the OU that will store the users.  If more then one OU is found with
'this name it will prompt you.  If left blank the "Users" container will be used
strOU = "ScriptTest" 

'Optional Entries - Leave the property blank if not needed i.e. ""
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
strHomeFolderRoot = "\\starbase\test" 'UNC path to home folder root
strHomeDrive = "H:" 'Drive letter for home folder
strProfilePath = "" 'UNC path to profile location
strScript = "student.bat" 'Logon script
strDescription = "Imported User" 'Users description

'Add the users to the following groups, separate each group with a comma.
strGroups = "Home Users"

'Password Settings
'~~~~~~~~~~~~~~~~~
'This script can generate a password for you automatically.  To do so set
'intSetPasswordLengh to the number of characters you want the random password to be.
'If you want all users to have the same password set strPassword to the password.
intPasswordLength = 8
strPassword = ""

'Folder Settings
'~~~~~~~~~~~~~~~
'Set the permissions on the home folders.  Put the string in the following format
'"group1:f,group2:c" permissions you can use = r=Read Only; w=Write; c=Change; f=Full
strFolderPermissions = "Domain Admins:f"

'Set the owners permission using the same permissions from above
strOwnerPermission = "c"

'*****************************************************************************************
'         !!!!!!!!!!!!!! DO NOT CHANGE ANYTHING BELOW THIS LINE !!!!!!!!!!!!!!
'*****************************************************************************************

'Scan the import file for errors
CheckInputFile

'Create the log file and get domain information
CreateLogFile
GetDomainInfo

'Create dictionary objects and OU object
Set objVerifiedGroups = GetGroups 'Verifies that the groups exist
Set objPerms = GetFolderPermissions 'Verifies that the groups exit and perms are correct
Set objOU = GetOU

'Initialize counters
intNewUserCount = 0
intNewFolderCount = 0

'Look for any errors before the script starts to create accounts
If Err Then
   MsgBox "At this point no accounts have been created but their was an error " & _
      "detected somewhere.  The script will now exit." & vbCRLF & "Error Code: " & _
      Err.Number & vbCRLF & "Error Description: " & Err.Description,vbCritical, _
      "Error Detected"
   Err.Clear
   WScript.Quit
End If

'Step through the import file
While txtSourceCSV.AtEndOfLine = False
   
   'Get the new users name from the file and break it up into first and last name
   strImportedData = txtSourceCSV.ReadLine
   ImportUserName
   
   'Since we need to have a unique username this section will loop thru generating a
   'username using the first letter of the first name and attempt to create the
   'account.  If that fails it will then attempt to create an account using the first
   'two letters from the first name, and so on until an account is created
   For intIndex = 1 to Len(strFirstName)
     
      'Generate a Random Password if requested.
      If intPasswordLength <> 0 Then
         strPassword = GeneratePassword
      End If
      
      'Attempt to create the account
      Set objNewUser = CreateUser
      
      'If there is an error then clear it and make sure all the letters from the first
      'name haven't been used.  If there wasn't an error then the account was created
      'and the script will continue to create the home folder and add to groups
      If Err Then
         Err.Clear
         If intIndex = Len(strFirstName) Then
            If bolReportOnly Then
               strError = ",<--- Problem with this user,Account will not be created"
               strOutput = """" & objNewUser.DisplayName & """," & ","  & strError               
            Else
               strError = ",<--- Problem with this account,Account not created"
               strOutput = """" & objNewUser.DisplayName & """," & ","  & strError               
            End If
            Exit For
         End If
      Else
         intNewUserCount = intNewUserCount + 1
         
         'Now that the account has been made we will make the users home folder
         If bolReportOnly Then
            intNewFolderCount = intNewFolderCount + 1
         
            'Write the users information to the log file
            strOutput = """" & objNewUser.DisplayName & """," & _
            objNewUser.SamAccountName & strError
            
         Else
            If strHomeFolderRoot <> "" Or strHomeDrive <> "" Then
               MakeHomeFolder   
            Else 'No folder was requested
               strError = strError & ",No home folder created"  
            End If         
            Err.Clear
            
            'Add the user to the appropriate groups
            AddtoGroups
            
            'Write the users information to the log file
            strOutput = """" & objNewUser.DisplayName & """," & _
            objNewUser.SamAccountName & "," & strPassword & strError            
         End If
         
         'Exit the loop if the user was created
         Exit For
      End If 
   Next
   txtLogFile.WriteLine(strOutput) 
Wend
   
'Close the CSV and Log files
txtSourceCSV.Close
txtLogFile.Close

'Display a message to the user when complete
If bolReportOnly Then
   MsgBox "The import report has been created.  No accounts were imported" & vbCRLF & _
   "Report Location: " & strLogFile & vbCRLF & intNewUserCount & " users would be created..." & _
   vbCRLF & intNewFolderCount & " folders would be created...",vbOkOnly,"Report Complete"
Else
   MsgBox "The users have been imported.  View the log for more information." & vbCRLF & _
   "Log File Location: " & strLogFile & vbCRLF & intNewUserCount & " users created..." & _
   vbCRLF & intNewFolderCount & " folders created...",vbOkOnly,"Import Complete"
End If

'Close objects
Set objPerms = Nothing
Set objVerifiedGroups = Nothing
Set objOU = Nothing
Set txtSourceCSV = Nothing
Set txtLogFile = Nothing

'*****************************************************************************************
Sub CheckInputFile
   'This Sub will check the input file for any errors, if found it will
   'let you know what line is bad and exit the script.
   
   On Error Resume Next   
  
   Dim objRegExp, objFSO, intLineCount, intCommaCount

   Set objRegExp = New RegExp
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set txtSourceCSV = objFSO.OpenTextFile(strSourceCSV)
   
   'Exit if there is an error opening the CSV file
   If Err Then
      MsgBox "Cannot find CSV file.  The script will now exit.",vbCritical,"Missing CSV"
      Err.Clear
      Wscript.Quit
   End If
   
   objRegExp.Pattern = ","
   intLineCount = 0
   
   While txtSourceCSV.AtEndOfLine = False
      strImportedData = txtSourceCSV.ReadLine
      intCommaCount = 0
      intLineCount = intLineCount + 1
      
      'Scan each line for comma's
      For intIndex = 1 to Len(strImportedData)
         If Mid(strImportedData,intIndex,1) = "," Then
            intCommaCount = intCommaCount + 1
         End If
      Next
      
      'Check for comma errors 
      Select Case intCommaCount
         Case 0
            MsgBox "Bad format in line " & intLineCount & ".  Missing a comma.", _
               vbCritical,"Error in CSV File"
            WScript.Quit
         Case 1
         Case Else
            MsgBox "Bad format in line " & intLineCount & ".  Too many commas.", _
               vbCritical,"Error in CSV File"
            WScript.Quit
      End Select
      
      'Look for odd characters in each line, report any problems 
      For intIndex = 33 to 64
         If inStr(strImportedData,Chr(intIndex)) Then
            Select Case intIndex
               Case 39 'Skip 39 = apostrophe 
               Case 44 'Skip 44 = comma
               Case 45 'Skip 45 = dash
               Case 46 'Skip 46 = period
               Case Else
                  MsgBox "Bad format in line " & intLineCount & ".  Line " & _
                     "contains a """ & Chr(intIndex) & """",vbCritical, _
                     "Error in CSV File"
                  WScript.Quit
            End Select
         End If            
      Next
   Wend
   
   'Close the source CSV file
   txtSourceCSV.Close
   
   'Close objects
   Set objRegExp = Nothing
   Set objFSO = Nothing
   Set txtSourceCSV = Nothing
   
End Sub
'*****************************************************************************************
Sub CreateLogFile

   On Error Resume Next
   
   Dim objFSO
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")

   'Add the CSV extension to the log file if missing or incorrect
   If UCase(Right(strLogFile,4)) <> ".CSV" Then
      If InStrRev(strLogFile,".") Then
         strLogFile = Left(strLogFile,Len(strLogFile) - 4) & ".csv"
      Else
         strLogFile = strLogFile & ".csv"
      End If
   End If
   
   'Create the log file and open the source file
   Set txtLogFile = objFSO.CreateTextFile(strLogFile)
   Set txtSourceCSV = objFSO.OpenTextFile(strSourceCSV)
   
   'Exit if there is an error creating the log file
   If Err Then
      MsgBox "Cannot create the log file.  The script will now exit.",vbCritical,"Missing CSV"
      Err.Clear
      Wscript.Quit
   End If
   
   'Create the log file heading
   If bolReportOnly Then
      txtLogFile.WriteLine("User Import Report,Report Completed On " & Date)
      txtLogFile.WriteLine("THIS IS A REPORT ONLY,ACTIVE DIRECTORY WAS NOT MODIFIED")
      txtLogFile.WriteLine("Display Name,UserName,Errors")
   Else
      txtLogFile.WriteLine("User Import Log, Import Completed On " & Date)
      txtLogFile.WriteLine("Display Name,UserName,Password,Errors")   
   End If
   
   'Prepare Password Settings
   If strPassword <> "" Then
      intPasswordLength = 0
   End If
   
   Set objFSO = Nothing
   
End Sub
'*****************************************************************************************
Sub GetDomainInfo
   'Returns domain information
   
   On Error Resume Next
   
   Dim objNet, objRootDSE, objRegExp  
   
   Set objNet = CreateObject("WScript.Network") 'Used to get the domain name
   Set objRootDSE = GetObject("LDAP://rootDSE") 'Used to connect to the domain
   Set objRegExp = New RegExp 'Used for String manipulation and checking
   
   'Exit the script if it can't get the domain information
   If Err Then
      MsgBox "You are either not on a domain or not connected to the network.  " & _
         "The script will now exit.",vbCritical,"Network Missing"
      Err.Clear
      WScript.Quit
   End If
   
   strDefaultNamingContext = objRootDSE.Get("DefaultNamingContext")
   strNetBIOSDomain = objNet.UserDomain 'Get the NetBIOS Domain Name
   
   'Get the Domain Name for the User Principle Name i.e. domain.com
   objRegExp.Global = True
   objRegExp.Pattern = ",DC="
   strDomain = objRegExp.Replace(UCase(strDefaultNamingContext),".")
   strDomain = Right(LCase(strDomain),Len(strDomain) - 3)
   objRegExp.Global = False
   
   'Close objects
   Set objNet = Nothing
   Set objRootDSE = Nothing
   Set objRegExp = Nothing
   
End Sub 
'*****************************************************************************************
Function GetGroups
   'Now lets get the information on the groups we are going to add the users to
   
   On Error Resume Next   
   
   Dim intIndex, strTempGroup, objGroups, Key, objADHelper, objGroup
   
   Set objGroups = CreateObject("Scripting.Dictionary")
   Set GetGroups = CreateObject("Scripting.Dictionary")
   Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects

   'Exit if the ADHelper object isn't installed
   If Err Then
      MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
      "The script will now exit.",vbCritical,"Network Missing"
      Err.Clear
      WScript.Quit
   End If
   
   If strGroups <> "" Then
      'Add each item from the comma separated variable to a Dictionary Object
      For intIndex = 1 to Len(strGroups)
         If Mid(strGroups,intIndex,1) <> "," Then
            strTempGroup = strTempGroup & Mid(strGroups,intIndex,1)
         Else
            objGroups.Add strTempGroup,""
            strTempGroup = ""
         End If   
      Next
      objGroups.Add strTempGroup,""
      
      'Look thru each item in the Dictionary Object see if the group exists in
      'Active Directory.  If not then exit the script.
      For Each Key in objGroups
         Set objGroup = objADHelper.GroupObject(Key)
         GetGroups.Add objGroup.DistinguishedName, ""
         If Err Then
            MsgBox "Cannot find the group " & Key & ".  The script will now exit.", _
            vbCritical,"Invalid Group"
            Err.Clear
            Wscript.Quit
         End If       
      Next
   End If
   
   Set objGroups = Nothing
   Set objGroup = Nothing
   Set objADHelper = Nothing
   
End Function
'*****************************************************************************************
Function GetFolderPermissions
   'Pull the information from the Folder Permissions String and add it to a 
   'Dictionary Object
   
   On Error Resume Next
   
   Dim objADHelper, strTemp, strTempPerm, strTempGroup, Key
   
   Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects

   'Exit if the ADHelper object isn't installed
   If Err Then
      MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
      "The script will now exit.",vbCritical,"Network Missing"
      Err.Clear
      WScript.Quit
   End If
   
   If strFolderPermissions <> "" Then
      If Right(strFolderPermissions,1) <> "," Then
         strFolderPermissions = strFolderPermissions & ","
      End If
      Set GetFolderPermissions = CreateObject("Scripting.Dictionary")
      For intIndex = 1 to Len(strFolderPermissions)
         
         Select Case Mid(strFolderPermissions,intIndex,1)
            Case ":"
               strTempGroup = strTemp
               strTemp = ""
            Case ","
               strTempPerm = strTemp
               GetFolderPermissions.Add strTempGroup,strTempPerm
               strTemp = ""
            Case Else
               strTemp = strTemp & Mid(strFolderPermissions,intIndex,1)
         End Select
      Next
      
      For Each Key in GetFolderPermissions
         Set objGroup = objADHelper.GroupObject(Key)
         If Err Then
            MsgBox "Cannot find the group " & Key & ".  The script will now exit.", _
               vbCritical,"Invalid Group"
            Err.Clear
            Wscript.Quit
         End If 
         CheckPermissions(GetFolderPermissions.Item(Key))
      Next
   End If
   
   CheckPermissions(strOwnerPermission)
   
   Set objGroup = Nothing
   Set objADHelper = Nothing
   
End Function
'*****************************************************************************************
Sub CheckPermissions(strPermission)
   Select Case UCase(strPermission)
   Case "F" 'Full Control
   Case "C" 'Change Control
   Case "W" 'Write Control
   Case "R" 'Read Only Control
   Case Else
      MsgBox """" & strPermission & """ is not a valid permission " & _
      "setting for new folders.  You must choose one of the following:" & vbCRLF & _  
      "r=Read Only; w=Write; c=Change; f=Full" & vbCRLF & "The script will now " & _
      "exit.",vbCritical,"Invalid Permission Setting"
      WScript.Quit
   End Select
End Sub
'*****************************************************************************************
Function GetOU
   'If the OU is blank then set it up to place users in the Users container
   
   On Error Resume Next
   
   Dim objADHelper
   
   Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects

   'Exit if the ADHelper object isn't installed
   If Err Then
      MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
      "The script will now exit.",vbCritical,"Network Missing"
      Err.Clear
      WScript.Quit
   End If
   
   If strOU = "" Then
      Set GetOU = GetObject("LDAP://CN=Users," & strDefaultNamingContext)
   Else
      Set GetOU = objADHelper.OUObject(strOU)
   End If
   
   If Err Then
      MsgBox """" & strOU & """ is not a valid OU.  The script will now exit.", _
      vbCritical,"Invalid OU"
      WScript.Quit
   End If
   
   Set objADHelper = Nothing
   
End Function
'*****************************************************************************************
Sub ImportUserName

   'Get the last and first names from the input file and put check for initials or
   'or name suffixes.

   Dim intCommaPosition, objRegExp, intSpacePosition, strTempName, intCharPosition
   Dim intIndex2, bolCharFound

   On Error Resume Next

   Set objRegExp = New RegExp 'Used for String manipulation and checking
   
   'Set variables to an empty string
   strNameSuffix = ""
   strError = ""   
   strInitial = ""
   bolCharFound = False
   
   'Look for the comma and split the variable into the first and last name
   intCommaPosition = inStr(strImportedData,",")
   strLastName = Trim(Left(strImportedData,intCommaPosition - 1))
   strFirstName = Trim(Right(strImportedData,Len(strImportedData) - intCommaPosition))
   
   'Look for a space in the last name, if there is, mark it in the log file. Then it tries
   'to determine if the second part is a suffix or a valid last name.
   objRegExp.Pattern = " "
   If objRegExp.Test(strLastName) Then
      
      'If there is a space found flag it as a possible error.  The script will attempt to
      'handle the space properly.
      If bolReportOnly Then
         strError = strError & ",<--- This user may not import properly"
      Else
         strError = strError & ",<--- Verify This Account"
      End If
      
      'Try to determine if the second word is a name suffix
      intSpacePosition = inStrRev(strLastName," ")
      If Len(Trim(Right(strLastName,Len(strLastName) - intSpacePosition))) < 4 Then
         strNameSuffix = UCase(Trim(Right(strLastName,Len(strLastName) - intSpacePosition)))
         If StrNameSuffix = "JR" or StrNameSuffix = "JR." Then
            StrNameSuffix = "Jr."
         End If
         If StrNameSuffix = "SR" or StrNameSuffix = "SR." Then
            StrNameSuffix = "Sr."
         End If
         strLastName = Trim(Left(strLastName,intSpacePosition))
      End If
      
      'Fix the case of the last name if there is a space in it
      intSpacePosition = inStrRev(strLastName," ")
      If intSpacePosition <> 0 Then
         strTempName = UCase(Trim(Left(strLastName,intSpacePosition)))
         strLastName = UCase(Trim(Right(strLastName,Len(strLastName) - intSpacePosition)))
         strLastName = UCase(Left(strTempName,1)) & LCase(Right(strTempName,Len(strTempName) - 1)) & _
         " " & UCase(Left(strLastName,1)) & LCase(Right(strLastName,Len(strLastName) - 1))
      Else
         strLastName = UCase(Left(strLastName,1)) & LCase(Right(strLastName,Len(strLastName) - 1))
      End If
   Else
      strLastName = UCase(Left(strLastName,1)) & LCase(Right(strLastName,Len(strLastName) - 1))      
   End If

   'Fix the case of the last name if their is an odd character in it
   For intIndex2 = 38 to 45
      intCharPosition = inStrRev(strLastName,Chr(intIndex2))
      If intCharPosition <> 0 Then
         bolCharFound = True
         strTempName = UCase(Trim(Left(strLastName,intCharPosition)))
         strLastName = UCase(Trim(Right(strLastName,Len(strLastName) - intCharPosition)))
         strLastName = UCase(Left(strTempName,1)) & LCase(Right(strTempName,Len(strTempName) - 1)) & _
         UCase(Left(strLastName,1)) & LCase(Right(strLastName,Len(strLastName) - 1))
      End If
   Next
   
   If bolCharFound Then
      strFirstName = UCase(Left(strFirstName,1)) & LCase(Right(strFirstName,Len(strFirstName) - 1))
   End If 
   
   'Look for a space in the first name if there is, mark it in the log file. Then it tries
   'to determine if the second part is a middle initial or a valid first name
   If objRegExp.Test(strFirstName) Then
      If InStr(strError,",") = 0 And strError <> "" Then
      
         'If there is a space found flag it as a possible error.  The script will attempt to
         'handle the space properly.
         If bolReportOnly Then
            strError = strError & ",<--- This user may not import properly"
         Else
            strError = strError & ",<--- Verify This Account"
         End If
      End If
      
      'Try to determine if the second word is a middle initial
      intSpacePosition = inStrRev(strFirstName," ")
      If Len(Trim(Right(strFirstName,Len(strFirstName) - intSpacePosition))) < 3 Then
         strInitial = UCase(Trim(Right(strFirstName,Len(strFirstName) - intSpacePosition)))
         strFirstName = Trim(Left(strFirstName,intSpacePosition))
      End If
      
      'Fix the case of the first name if their is a space in it
      intSpacePosition = inStrRev(strFirstName," ")
      If intSpacePosition <> 0 Then
         strTempName = UCase(Trim(Left(strFirstName,intSpacePosition)))
         strFirstName = UCase(Trim(Right(strFirstName,Len(strFirstName) - intSpacePosition)))
         strFirstName = UCase(Left(strTempName,1)) & LCase(Right(strTempName,Len(strTempName) - 1)) & _
         " " & UCase(Left(strFirstName,1)) & LCase(Right(strFirstName,Len(strFirstName) - 1))
      Else
         strFirstName = UCase(Left(strFirstName,1)) & LCase(Right(strFirstName,Len(strFirstName) - 1))
      End If
   Else
      strFirstName = UCase(Left(strFirstName,1)) & LCase(Right(strFirstName,Len(strFirstName) - 1))      
   End If
   
   Set objRegExp = Nothing
   
End Sub
'*****************************************************************************************
Function GeneratePassword
   'Generate a random password for each user
   
   Dim intIndex, intRandomNumber
   
   Randomize Timer
   
   For intIndex = 1 to intPasswordLength
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
Function CreateUser
   'Put the username in the correct format
   
   Dim strDisplayName, strDistinguishedName, strUserName, objRegExp, intIndex2
   
   Set objRegExp = New RegExp 'Used for String manipulation and checking

   objRegExp.Pattern = " "
   If Not bolLastNameFirst Then
      strUserName = Left(strFirstName,intIndex) & objRegExp.Replace(strLastName,"")
   Else
      strUserName = objRegExp.Replace(strLastName,"") & Left(strFirstName,intIndex)
   End If
   
   'Fix the username by removing the folowing characters !"#$%&'()*+`./ 0-9 :;<=>?@
   For intIndex2 = 33 to 64
      If inStr(strUserName,Chr(intIndex2)) And intIndex2 <> 45 Then
         objRegExp.Pattern = "\" & Chr(intIndex2)
         strUserName = objRegExp.Replace(strUserName,"")
      End If
   Next
          
   strDisplayName = Trim(strLastName & ", " & strFirstName & " " & strNameSuffix)      
   strDistinguishedName = Trim(strLastName & "\, " & strFirstName & " " & strNameSuffix)      
   
   'Create the user and set the properties
   Set CreateUser = objOU.Create("User", "cn=" & strDistinguishedName)
   CreateUser.Put "SAMAccountName", strUserName
   CreateUser.Put "SN", strLastName
   CreateUser.Put "GivenName", strFirstName
   CreateUser.Put "DisplayName", strDisplayName
   CreateUser.Put "UserPrincipalName", strUserName & "@" & strDomain
   CreateUser.Put "UserAccountControl", 544 '544 = normal account, no password required
   
   If strInitial <> "" Then
      CreateUser.Put "Initials", Left(strInitial,1)
   End If
   
   If strScript <> "" Then
      CreateUser.Put "ScriptPath", strScript
   End If
   
   If strHomeFolderRoot <> "" Or strHomeDrive <> "" Then
      CreateUser.Put "HomeDirectory", strHomeFolderRoot & "\" & CreateUser.SamAccountName
      CreateUser.Put "HomeDrive", strHomeDrive
   End If
   
   If strProfilePath <> "" Then
      CreateUser.Put "ProfilePath", strProfilePath
   End If
   
   If strDescription <> "" Then
      CreateUser.Put "Description", strDescription
   End If

   If Not bolReportOnly Then
      CreateUser.SetInfo
      CreateUser.Setpassword(strPassword)
      CreateUser.Put "UserAccountControl", 512
      CreateUser.Put "pwdLastSet", 0 'User must change password at next login
      CreateUser.SetInfo
   End If
   
   Set objRegExp = Nothing
   
End Function
'*****************************************************************************************
Sub MakeHomeFolder
   'Remove any "\\"'s from the path, then but it back on the beginning.  This
   'allows the users to type in the path name with or without a "\" at the end.
   
   Dim Key, objRegExp, objShell, objFSO, strCMD, intCaclsError, strNewHomeFolder
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")   
   Set objRegExp = New RegExp 'Used for String manipulation and checking
   Set objShell = CreateObject("Wscript.Shell") 'Used to run the CACLS command
   
   objRegExp.Pattern = "\\\\"
   objRegExp.Global = True
   strNewHomeFolder = strHomeFolderRoot & "\" & objNewUser.SamAccountName
   strNewHomeFolder = objRegExp.Replace(strNewHomeFolder,"\")            
   If Left(strNewHomeFolder,1) = "\" Then
      strNewHomeFolder = "\" & strNewHomeFolder
   End If            
   objRegExp.Global = False
   
   If Not objFSO.FolderExists(strNewHomeFolder) Then
               
      'Create the folder
      objFSO.CreateFolder(strNewHomeFolder)
      
      If Not objFSO.FolderExists(strNewHomeFolder) Then
         strError = strError & ",Error creating home folder"
      Else 'Everything was ok
         intNewFolderCount = intNewFolderCount + 1
      End If
      
      'Add the user to the Dictionary Object so they will have permissions on 
      'their folder
      objPerms.Add objNewUser.SamAccountName, strOwnerPermission
      
      'Build the CACLS command using the Dictionary Object and execute it, record
      'any errors to the intCaclsError variable
      strCMD = "cmd /c echo y| cacls " & """" & strNewHomeFolder & """ /c /t /g "
      For Each Key in objPerms
         strCMD = strCMD & """" & strNetBIOSDomain & "\" & Key & """" & ":" & _
            objPerms.Item(Key) & " "  
      Next              
      intCaclsError = objShell.Run(strCMD,0,true)
      
      'Remove the user from the Dictionary Object 
      objPerms.Remove(objNewUser.SamAccountName)
      
      'If the CACLS command fails then remove the folder and write an error to
      'the log.
      If intCaclsError <> 0 AND objFSO.FolderExists(strNewHomeFolder) Then
         objFSO.DeleteFolder(strNewHomeFolder)
         strError = strError & ",Could not modify permissions"
         intNewFolderCount = intNewFolderCount - 1
      End If
   Else 'Folder already existed
      strError = strError & ",Folder already existed,permissions not changed"
   End If
   
   Set objFSO = Nothing
   Set objRegExp = Nothing
   Set objShell = Nothing
   
End Sub
'*****************************************************************************************
Sub AddtoGroups

   Const ADS_PROPERTY_APPEND = 3
   
   Dim Key, strADSPath
         
   'Now we will add the user to the requested groups         
   For Each Key in objVerifiedGroups
            
      'Build the Group path and create the Group object
      strADSPath = "LDAP://" & Key
      Set objGroup = GetObject(strADSPath) 
      
      'Add the user the the appropriate groups                            
      objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(objNewUser.DistinguishedName)
      objGroup.SetInfo
   Next
End Sub 
'*****************************************************************************************