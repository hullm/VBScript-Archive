'Created by Matthew Hull on 8/16/04

'This script will set the logon hours for each user in an OU using the Net User
'command.

Option Explicit

On Error Resume Next

Dim strOU, objADHelper, objOU, objUser, strCMD, intNetUserError, objShell
Dim strLogonHours

'**************************************************************************
'All user input is done in this section

strOU = "Script Test"
strLogonHours = "M-F,8am-3pm"

'The following comments are from http://www.jsiinc.com/SUBP/tip7500/rh7540.htm
' - Days can be spelled out (for example, Monday) or abbreviated 
'   (for example, M,T,W,Th,F,Sa,Su).
' - Hours can be in 12-hour notation (1PM or 1P.M.) or 24-hour notation (13:00).
' - A value of blank means that the user can never log on.
' - A value of all means that a user can always log on.
' - Use a hyphen (-) to mark a range of days or times. For example, to create a range 
'   from Monday through Friday, type either M-F, or monday-friday. To create a range of time 
'   from 8:00 P.M. to 5:00 P.M., type 8:00am-5:00pm, 8am-5pm, or 8:00-17:00.
' - Separate the day and time items with commas (for example, monday,8am-5pm).
' - Separate day and time units with semicolons 
'   (for example, monday,8am-5pm;tuesday,8am-4pm;wednesday,8am-3pm).
' - Do not use spaces between days or times.
'**************************************************************************

'Create the ADHelper and Shell objects
Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects
Set objShell = Wscript.CreateObject("Wscript.Shell")

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

'Loop thru each user in the OU
For Each objUser in objOU

   'Verify the user is of type user
   If objUser.Class = "user" Then
   
      'Build the Net User command
      strCMD = "cmd /c Net User " & objUser.SamAccountName & " /time:" & _
      strLogonHours
      
      'Run the Net User comamnd, return any errors to intNetUserError
      intNetUserError = objShell.Run(strCMD,0,true) 
   
   End If
Next

MsgBox "The users logon times have been set",vbOkOnly,"Complete"

'Close open objects
Set objADHelper = Nothing
Set objOU = Nothing
Set objUser = Nothing