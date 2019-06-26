'Created by Matthew Hull on 12/18/03 Version 2.0
'Documented on 4/25/04

'This script will uninstall printers on a local workstation.  The way it works is you enter
'the UNC names of the printers into a Dictionary Object then the script loops thru the 
'object and uninstalls each printer.  If there is an error it will record a warning 
'in the event viewer.  

Option Explicit

Dim objWSHNetwork, objWSHShell, objDict, strPrinter

'Create a Dictoanry object, this will hold all the printers you want to install
Set objDict = CreateObject("Scripting.Dictionary")

'Create a Network object, this will be used to add the printer
Set objWSHNetwork = CreateObject("WScript.Network")

'Create a Shell object, this will be used to write any errors to the event viewer
Set objWSHShell = CreateObject("WScript.Shell")

'******************************************************************************************
'To remove a printer type objDict.Add "\\Server\Share", ""  If you want the printer to be
'the default printer put the word Default between the last two quotes.

objDict.Add "\\server\share1", ""
objDict.Add "\\server\share2", ""
'******************************************************************************************

On Error Resume Next

'Loop thru each printer in the Dictionary object
For Each strPrinter In objDict

   'Remove the printer from the computer
   objWSHNetwork.RemovePrinterConnection strPrinter, True, True
   
   'If their was an error removing the printer then write a message to the Event Viewer 
   If Err Then
      objWSHShell.LogEvent 2, "There was an error uninstalling " & strPrinter & " " & _
      Err.Description
      Err.Clear
   End If
Next