'Created by Matthew Hull on 8/9/04

'This script will display the name of your computer

Option Explicit

On Error Resume NExt

Dim objShell, strCompName

'Create the Shell Object, this will be used to get environmental variables
Set objShell = CreateObject("WScript.Shell")

'Get the logon server name
strCompName = objShell.ExpandEnvironmentStrings("%ComputerName%")

'Display the message to the user
MsgBox "Your computer name is " & strCompName,vbOkOnly,"Computer Name"

'Close open variables
Set objShell = Nothing