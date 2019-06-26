'Created by Matthew Hull on 8/04/04
'Last updated 8/05/04

'This script will modify the terminal services settings for all users in an OU.  It
'uses a program called TSCMD from www.systemtools.com.  Please view the readme file
'for more information on the settings.

Option Explicit

On Error Resume Next

Dim strOU, strInitialProgram, strWorkingDirectory, strInheritInitialProgram 
Dim strAllowLogonTerminalServer, strTimeoutConnection, strTimeoutDisconnect
Dim strTimeoutIdle, strDeviceClientDrives, strDeviceClientPrinters
Dim strDeviceClientDefaultPrinter, strBrokenTimeoutSettings, strReconnectSettings
Dim strModemCallbackSettings, strModemCallbackPhoneNumber, strShadowingSettings 
Dim strTerminalServerProfilePath, strTerminalServerHomeDir, strTerminalServerHomeDirDrive
Dim objShell, objADHelper, objOU, objRegExp, strServer, objUser

'***************************************************************************************
'All user entry is done in this section

strOU = "ScriptTest"
strInitialProgram = ""
strWorkingDirectory = ""
strInheritInitialProgram  = ""
strAllowLogonTerminalServer = ""
strTimeoutConnection = ""
strTimeoutDisconnect = ""
strTimeoutIdle = ""
strDeviceClientDrives = ""
strDeviceClientPrinters = ""
strDeviceClientDefaultPrinter = ""
strBrokenTimeoutSettings = ""
strReconnectSettings = ""
strModemCallbackSettings = ""
strModemCallbackPhoneNumber = ""
strShadowingSettings = ""
strTerminalServerProfilePath = ""
strTerminalServerHomeDir = ""
strTerminalServerHomeDirDrive = ""

'***************************************************************************************

'Create the Shell Object, this will be used to get environmental variables
Set objShell = CreateObject("WScript.Shell")

'Create the ADHelper object
Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects

'Exit if the ADHelper object isn't installed
If Err Then
   MsgBox "You must have ADHelper 1.0 or later installed on your PC.  " & _
   "The script will now exit.",vbCritical,"Network Missing"
   Err.Clear
   WScript.Quit
End If

'Create an OU object using the ADHelper object
Set objOU = objADHelper.OUObject(strOU)

'Exit the script if the OU isn't found
If Err Then
   MsgBox """" & strOU & """ is not a valid OU.  The script will now exit.", _
   vbCritical,"Invalid OU"
   WScript.Quit
End If

'Create a Regular Expression Object
Set objRegExp = New RegExp

'Set the pattern for the regular expression
objRegExp.Pattern = "\\\\"

'Get the logon server name
strServer = objShell.ExpandEnvironmentStrings("%logonserver%")
strServer = lcase(objRegExp.Replace(strServer,""))

'Loop thru each user in the OU
For Each objUser in objOU
   
   'Verify the user is of type user
   If objUser.Class = "user" Then
      
      If strInitialProgram <> "" Then
         Call RunTSCMD("InitialProgram",strInitialProgram)
      End If
      
      If strWorkingDirectory <> "" Then   
         Call RunTSCMD("WorkingDirectory",strWorkingDirectory)
      End If         
      
      If strInheritInitialProgram <> "" Then
         Call RunTSCMD("InheritInitialProgram",strInheritInitialProgram)
      End If
         
      If strAllowLogonTerminalServer <> "" Then      
         Call RunTSCMD("AllowLogonTerminalServer",strAllowLogonTerminalServer)
      End If   
         
      If strTimeoutConnection <> "" Then   
         Call RunTSCMD("TimeoutConnection",strTimeoutConnection)
      End If   
         
      If strTimeoutDisconnect <> "" Then   
         Call RunTSCMD("TimeoutDisconnect",strTimeoutDisconnect)
      End If   
         
      If strTimeoutIdle <> "" Then   
         Call RunTSCMD("TimeoutIdle",strTimeoutIdle)
      End If   
         
      If strDeviceClientDrives <> "" Then   
         Call RunTSCMD("DeviceClientDrives",strDeviceClientDrives)
      End If   
         
      If strDeviceClientPrinters <> "" Then   
         Call RunTSCMD("DeviceClientPrinters",strDeviceClientPrinters)
      End If   
        
      If strDeviceClientDefaultPrinter <> "" Then  
         Call RunTSCMD("DeviceClientDefaultPrinter",strDeviceClientDefaultPrinter)
      End If   
        
      If strBrokenTimeoutSettings <> "" Then   
         Call RunTSCMD("BrokenTimeoutSettings",strBrokenTimeoutSettings)
      End If   
         
      If strReconnectSettings <> "" Then   
         Call RunTSCMD("ReconnectSettings",strReconnectSettings)
      End If   
         
      If strModemCallbackSettings <> "" Then  
         Call RunTSCMD("ModemCallbackSettings",strModemCallbackSettings)
      End If   
        
      If strModemCallbackPhoneNumber <> "" Then   
         Call RunTSCMD("ModemCallbackPhoneNumber",strModemCallbackPhoneNumber)
      End If   
         
      If strShadowingSettings <> "" Then   
         Call RunTSCMD("ShadowingSettings",strShadowingSettings)
      End If   
         
      If strTerminalServerProfilePath <> "" Then   
         Call RunTSCMD("TerminalServerProfilePath",strTerminalServerProfilePath)
      End If   
         
      If strTerminalServerHomeDir <> "" Then   
         Call RunTSCMD("TerminalServerHomeDir",strTerminalServerHomeDir)
      End If   
         
      If strTerminalServerHomeDirDrive <> "" Then   
         Call RunTSCMD("TerminalServerHomeDirDrive",strTerminalServerHomeDirDrive)
      End If  
         
   End If
Next

MsgBox "The users have been modified",vbOkOnly,"Complete"

Sub RunTSCMD(strProperty, strSetting)

   On Error Resume Next
   
   Dim strCMD, intTSCMDError
   
   'Build the command
   strCMD = "tscmd " & strServer & " " & objUser.SamAccountName & " " & _
   strProperty & " " & strSetting
      
   'Execute the TSCMD Command
   intTSCMDError = objShell.Run(strCMD,0,true) 
   
   'Exit if an Error is found
   If intTSCMDError <> 0 Then
      MsgBox "TSCMD doesn't appear to be installed or it isn't located in a folder in the path", _
         vbCritical,"TSCMD Not Found"
      Wscript.Quit
   End If
   
End Sub