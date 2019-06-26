'Created by Matthew Hull on 8/9/04

'Deletes profiles on all computers in an OU.  This uses the DelProf program from
'Microsoft.

Option Explicit

On Error Resume Next

Dim strOU, objADHelper, objOU, objComputer, strCMDRoot, strCMD, objShell
Dim intDelProfError, strError, txtOUtPut, strErrorOutput, objFSO

'***************************************************************************************
strOU = "Lab300"
strErrorOutput = "C:\Delete Profile Error Log.txt"
'***************************************************************************************

'Create the ADHelper, Shell, and File System objects
Set objADHelper = CreateObject("ADHelper.wsc") 'Used to make the OU and Group objects
Set objShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

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

strCMDRoot = "delprof /q /i /c:\\"
strError = ""

'Loop thru each computer in the OU
For Each objComputer in objOU

   'Verify the object is of type computer
   If objComputer.Class = "computer" Then
      strCMD = strCMDRoot & objComputer.CN
      intDelProfError = objShell.Run(strCMD,1,true) 
      
      If Err Then
         MsgBox "DelProf is not installed on this computer.  The script will " & _
         "now exit",vbCritical,"Error"
         Err.Clear
         Wscript.Quit
      End IF
           
      If intDelProfError <> 0 Then
         strError = strError & objComputer.CN & vbCRLF
      End If      
   End If   
Next

'Display a message when complete
If strError = "" Then
   MsgBox "Profiles on remote computers deleted.",vbOkOnly,"Complete"
Else
   MsgBox "Profiles on remote computers deleted with the following exceptions." & _
   vbCRLF & strError,vbOkOnly,"Completed With Errors"
   Set txtOUtPut = objFSO.CreateTextFile(strErrorOutput)
   txtOutput.Write "Profiles on remote computers deleted with the following exceptions." & _
   vbCRLF & strError
   txtOutput.Close
   Set txtOUtPut = Nothing
End If

'Close open objects
Set objADHelper = Nothing
Set objOU = Nothing
Set objComputer = Nothing
Set objShell = Nothing
Set objFSO = Nothing