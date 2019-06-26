'Created by Matthew Hull 7/28/04

'This script will fix printers on a client after a servers name has changed.

Option Explicit

On Error Resume Next

Dim strOldServer, strNewServer, objWSHNetwork, colPrinters, intIndex, objRegExp
Dim objRemove, objInstall, strNewPrinter

'******************************************************************************************
'Set the old and new names of the server.  Some data needs to be modified in the Select
'statement if your sharenames and printer names don't match.

strOldServer = "Athena"
strNewServer = "GickDC"

'******************************************************************************************

'Create a Dictoanry object, this will hold all the printers you want to install
Set objRemove = CreateObject("Scripting.Dictionary")
Set objInstall = CreateObject("Scripting.Dictionary")

Set objRegExp = New RegExp
objRegExp.Pattern = UCase(strOldServer)

'Create a Network object, this will be used to add the printer
Set objWSHNetwork = CreateObject("WScript.Network")

'Get a collection of installed printers
Set colPrinters = objWSHNetwork.EnumPrinterConnections

'Loop through the returned array of printers
For intIndex = 1 to colPrinters.Count - 1 Step 2
   If objRegExp.Test(UCase(colPrinters(intIndex))) Then
      
      'If the share name and printer name don't match use this section to fix that
      Select Case UCase(colPrinters(intIndex))
         Case Ucase("\\" & strOldServer & "\" & "CY-4500N-PCL")
            strNewPrinter = "\\" & strNewServer & "\" & "Trish's4500N"
         Case Ucase("\\" & strOldServer & "\" & "CY-4500N-PS")
            strNewPrinter = "\\" & strNewServer & "\" & "SSS-4500PS"
         Case Ucase("\\" & strOldServer & "\" & "HP DeskJet 1600C NT4")
            strNewPrinter = "\\" & strNewServer & "\" & "LTHP1600C"   
         Case Else
            strNewPrinter = objRegExp.Replace(UCase(colPrinters(intIndex)),strNewServer)
      End Select
      
      'Build the dictionary object with the printers that need to be removed and added
      objRemove.Add colPrinters(intIndex),""
      objInstall.Add strNewPrinter,""
   End If
Next

'Remove the old and add the new
RemovePrinter(objRemove)
InstallPrinter(objInstall)

'Display a message when done
MsgBox "Printers Fixed",vbOkOnly,"Complete"

'Close objects
Set objWSHNetwork = Nothing
Set objRegExp = Nothing
Set objRemove = Nothing
Set objInstall = Nothing

'******************************************************************************************
Sub InstallPrinter(objDict)

   Dim objWSHNetwork, objWSHShell, strPrinter
   
   On Error Resume Next
   
   'Create a Network object, this will be used to add the printer
   Set objWSHNetwork = CreateObject("WScript.Network")
   
   'Create a Shell object, this will be used to write any errors to the event viewer
   Set objWSHShell = CreateObject("WScript.Shell")
   
   'Loop thru each printer in the Dictionary object
   For Each strPrinter in objDict
      
      'Add the printer to the computer
      objWshNetwork.AddwindowsPrinterConnection strPrinter   
   
      'If their was an error during the creation of the printer then log an event to the
      'Event Viewer
      If Err Then
         objWSHShell.LogEvent 2, "There was an error installing " & strPrinter & " " & _
         Err.Description
         Err.Clear
      End If
      
      'Check to see if the printer that was added should be the default, if so do it.
      If objDict.Item(strPrinter) = "Default" Then
         objWSHNetwork.SetDefaultPrinter strPrinter
      End If
   Next
   
   Set objWSWNetwork = Nothing
   Set objWSHShell = Nothing
   Set objDict = Nothing
   
End Sub
'******************************************************************************************
Sub RemovePrinter(objDict)
   Dim objWSHNetwork, objWSHShell, strPrinter
   
   On Error Resume Next
   
   'Create a Network object, this will be used to add the printer
   Set objWSHNetwork = CreateObject("WScript.Network")
   
   'Create a Shell object, this will be used to write any errors to the event viewer
   Set objWSHShell = CreateObject("WScript.Shell")
   
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
   
   Set objWSHNetwork = Nothing
   Set objWSHShell = Nothing
   Set objDict = Nothing
   
End Sub
'******************************************************************************************