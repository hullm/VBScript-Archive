'Created by Matthew Hull on 7/21/04

'This script will copy the users favorites to the users folder on the server.

Option Explicit

On Error Resume Next

Dim objShell, objFSO, objNet, strHomeShare, strUserProfile

'Create the shell and file system objects
Set objShell = WScript.CreateObject("WScript.Shell")
set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objNet = WScript.CreateObject("WScript.Network")
                                                    
'Use the Shell Object to retrieve system environment variables
strHomeShare = "\\Server\Share\" & objNet.UserName
strUserProfile = objShell.SpecialFolders("Favorites") 'Local

'Exit if the variables are blank
If strHomeShare = "" Or strUserProfile = "" Then
   objShell.LogEvent 2, "Cannot copy favorites to the server" & vbCRLF & _
   "Error Description: Variables not set"
   Wscript.Quit
End If

If objFSO.FolderExists(strHomeShare & "\Favorites") Then
   objFSO.DeleteFolder strHomeShare & "\Favorites"
End If

objFSO.CopyFolder strUserProfile,strHomeShare & "\Favorites"

If Err Then
   objShell.LogEvent 1, "Cannot copy favorites from " & strUserProfile & _
   " to " & strHomeShare & "\Favorites" & vbCRLF & "Error Description: " & _
   Err.Description
   Err.Clear
Else
   objShell.LogEvent 0, "Favorites successfully copied from " & _
   strUserProfile & " to " & strHomeShare & "\Favorites"   
End If

Set objShell = Nothing
Set objFSO = Nothing
Set objNet = Nothing