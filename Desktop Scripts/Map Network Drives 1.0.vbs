'Created by Matthew Hull on 12/13/04

'This script will map a drive on a local workstation.  The way it works is you enter
'the UNC names of the drive into a Dictionary Object then the script loops thru the 
'object and creates each drive.  If there is an error it will record a warning 
'in the event viewer.  

Option Explicit

Dim objWSHNetwork, objWSHShell, objDict, strDrive

'Create a Dictoanry object, this will hold all the drives you want to install
Set objDict = CreateObject("Scripting.Dictionary")

'Create a Network object, this will be used to add the drives
Set objWSHNetwork = CreateObject("WScript.Network")

'Create a Shell object, this will be used to write any errors to the event viewer
Set objWSHShell = CreateObject("WScript.Shell")

'******************************************************************************************
'To add a printer type objDict.Add "\\Server\Share", ""  If you want the printer to be
'the default printer put the word Default between the last two quotes.

objDict.Add "\\server\share1", "E:"
objDict.Add "\\server\share2", "F:"
'******************************************************************************************

On Error Resume Next

'Loop thru each drive in the Dictionary object
For Each strDrive in objDict
   
   'Add the drive to the computer
   objWshNetwork.MapNetworkDrive objDict.Item(strDrive), strDrive   
   
   'If their was an error during the creation of the drive then log an event to the
   'Event Viewer
   If Err Then
      objWSHShell.LogEvent 2, "There was an error installing " & strDrive & " " & _
      Err.Description
      Err.Clear
   End If
   
Next