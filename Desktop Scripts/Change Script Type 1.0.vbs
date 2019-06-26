'Created By Matthew Hull on some day a long time ago
'Documented on 4/25/04

Dim strScriptType,strCmd,obShell,intCMDError

'Create the Shell Object, this will be used to run the script command
Set objShell = WScript.CreateObject("WScript.Shell")

'Set the script type to nothing and start building the command
strScriptType = ""
strCmd = "cmd /c cscript /H:"

'Loop until the user enters a valid input
Do 
   strScriptType = InputBox("Please enter the script type you want: " & vbCRLF & _
   "1 for WScript" & vbCRLF & "2 for CScript","Choose Script Type")
   If strScriptType <> "1" And strScriptType <> "2" Then
      MsgBox "You must enter a 1 or a 2",vbOkOnly,"Bad Input"
      strScriptType = ""
   End If   
Loop While strScriptType = ""

'Check the users entry and set the Script type accordingly
If strScriptType = "1" then
   strScriptType = "WScript"
   Else
   strScriptType = "CScript"
End If

'Fininsh building the command
strCmd = strCmd & strScriptType

'Run the command to change the script type.  Return any errors to intCMDError
intCMDError = objShell.Run(strCmd,0,true)

'Display a message to the user when done.
MsgBox "Your scripting engine has been changed to " & strScriptType & ".",vbOkOnly,"Done"