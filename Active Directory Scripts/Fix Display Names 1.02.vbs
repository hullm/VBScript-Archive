'Created By Matthew Hull on some day a long time ago
'Documented on 4/21/04
'Last Updated 7/28/04

'This script is used to put the display name in last name first format.  This script is
'assuming that your users have the display name in first name first format and nothing
'for the first and last name fields.  It will generate the first and last name fields from
'the display name then change the display name format.  It doesn't care about jr's etc..
'Be Careful when running this script and try it in a test environment first.

Option Explicit

On Error Resume Next

Dim strLogFile, strOU, objADHelper, objOU, objRegExp, objFSO, txtLogFile, objUser
Dim strImportedData, intSpacePosition, strFirstName, strLastName, strError, strDistinguishedName
Dim StrNameSuffix, strInitial, strDisplayName, bolReportOnly, intSpacePositionRev

'*****************************************************************************************************
strLogFile = "C:\Fix Display Name.csv"
strOU = "ScriptTest"
bolReportOnly = True
'*****************************************************************************************************

'Create an OU Object
Set objADHelper = CreateObject("ADHelper.wsc")
Set objOU = objADHelper.OUObject(strOU)
'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
      "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

Set objRegExp = New RegExp
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set txtLogFile = objFSO.CreateTextFile(strLogFile)
If Err Then
   MsgBox "The log file is already open." & vbCRLF & "The script will now exit", _
      vbCritical,"Error"
   WScript.Quit
End If

objRegExp.Pattern = ","
objRegExp.Global = True

If bolReportOnly Then
   txtLogFile.WriteLine("THIS IS A REPORT ONLY.  ACTIVE DIRECTORY WAS NOT MODIFIED")
   txtLogFile.WriteLine("Report from " & objRegExp.Replace(objOU.DistinguishedName,";"))
Else
   txtLogFile.WriteLine("Fix Display Name log from " & objRegExp.Replace(objOU.DistinguishedName,";"))   
End If

txtLogFile.WriteLine("Display Name,User Name,First Name,Last Name,Status")

'Loop thru each user in the OU
For Each objUser in objOU

   'Verify the user is of type user
   If objUser.Class = "user" Then   
   
      'Get the current display name and split it into first name and last name
      strImportedData = objUser.DisplayName
      intSpacePositionRev = InStrRev(strImportedData," ")
      intSpacePosition = InStr(strImportedData," ")
 
      If intSpacePosition <> intSpacePositionRev Then
         intSpacePosition = intSpacePositionRev
      End If
      
      If intSpacePosition <> 0 Then
         strFirstName = Trim(Left(strImportedData,intSpacePosition - 1))
         strLastName = Trim(Right(strImportedData,Len(strImportedData) - intSpacePosition))
      End If
      objRegExp.Pattern = ","
      If Not objRegExp.Test(strImportedData) And intSpacePosition > 0 Then  
         
         'Look for a space in the last name.  If there is one attempt to fix it and alert the user
         objRegExp.Pattern = " "
         If objRegExp.Test(strLastName) Then
            If bolReportOnly Then
               strError = strError & "<--- This account may not convert properly"
            Else
               strError = strError & "<--- Verify This Account"
            End If
            intSpacePosition = inStrRev(strLastName," ")
            If Len(Trim(Right(strLastName,Len(strLastName) - intSpacePosition))) < 4 Then
               strNameSuffix = Trim(Right(strLastName,Len(strLastName) - intSpacePosition))
               If StrNameSuffix = "JR" or StrNameSuffix = "JR." Then
                  StrNameSuffix = "Jr."
               End If
               If StrNameSuffix = "SR" or StrNameSuffix = "SR." Then
                  StrNameSuffix = "Sr."
               End If
               strLastName = Trim(Left(strLastName,intSpacePosition))
            End If
         End If

         'Look for a space in the first name.  If there is one attempt to fix it and alert the user
         If objRegExp.Test(strFirstName) Then
            If strError = "" Then
               If bolReportOnly Then
                  strError = strError & "<--- This account may not convert properly"
               Else
                  strError = strError & "<--- Verify This Account"
               End If
            End If
            intSpacePosition = inStr(strFirstName," ")

            If intSpacePosition > 3 Then
               strInitial = Trim(Right(strFirstName,1))
               strFirstName = (Trim(Left(strFirstName,Len(strFirstName) - 1)))
            End If
         End If      
      
         'Put the first and last name into the proper case
         strLastName = UCase(Left(strLastName,1)) & LCase(Right(strLastName,Len(strLastName) - 1))
         strFirstName = UCase(Left(strFirstName,1)) & LCase(Right(strFirstName,Len(strFirstName) - 1))
         
         'Build the new display name
         strDisplayName = Trim(strLastName & ", " & strFirstName)          
         strDistinguishedName = Trim(strLastName & "\, " & strFirstName)
         
         If Not bolReportOnly Then
            objOU.MoveHere "LDAP://" & objUser.DistinguishedName, "cn=" & strDistinguishedName
            Set objUser = GetObject("LDAP://cn=" & strDistinguishedName & "," & objOU.DistinguishedName)               
         End If
                  
         'Write the value to the DisplayName,SN and GivenName Property
         objUser.Put "DisplayName", strDisplayName
         objUser.Put "SN", strLastName
         objUser.Put "GivenName", strFirstName
         
         'Write the new value to Active Directory
         If Not bolReportOnly Then
            objUser.SetInfo
         End If
         
         txtLogFile.WriteLine("""" & objUser.DisplayName & """," & objUser.SamAccountName & "," & _
            objUser.GivenName & "," & objUser.SN & "," & strError)
         strError = ""
      Else
         If bolReportOnly Then
            txtLogFile.WriteLine("""" & objUser.DisplayName & """," & objUser.SamAccountName & "," & _
               objUser.GivenName & "," & objUser.SN & ",Account Won't Be Modified")   
         Else
            txtLogFile.WriteLine("""" & objUser.DisplayName & """," & objUser.SamAccountName & "," & _
               objUser.GivenName & "," & objUser.SN & ",Account Not Modified")   
         End If
      End If       
   End If
Next

'Close the log file
txtLogFile.Close

'Display a message when done
If bolReportOnly Then
   MsgBox "Report created.  The report can be found here: " & strLogFile,vbOkOnly,"Report Created"
Else
   MsgBox "Display names fixed.  The log file can be found here: " & strLogFile,vbOkOnly,"Done"
End If

Set objADHelper = Nothing
Set objOU = Nothing
Set objRegExp = Nothing
Set objFSO = Nothing
Set txtLogFile = Nothing
Set objUSer = Nothing