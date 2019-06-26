'Created by Matthew Hull on 7/21/04
'Modified by Jason Purificato on 7/17/06

'This script will copy the users favorites to the users folder from the server. 

Option Explicit

On Error Resume Next

Dim objShell, objFSO, strHomeShare, strUserProfile

'Create the shell and file system objects
Set objShell = WScript.CreateObject("WScript.Shell")
set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

'Use the Shell Object to retrieve system environment variables
strHomeShare = objShell.ExpandEnvironmentStrings("%HOMESHARE%") 'Remote
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%") 'Local

'Exit if the variables are blank
If strHomeShare = "%HOMESHARE%" Or strUserProfile = "%USERPROFILE%" Then
   objShell.LogEvent 2, "Cannot copy favorites to the server" & vbCRLF & _
   "Error Description: Environmental variables not set"
   Wscript.Quit
End If

'If the Favorites folder exists on the server then backup the local favorites,
'delete them, then copy them from the server.
If objFSO.Exists(strHomeShare & "\Favorites") Then
   objFSO.copyFolder strUserProfile & "\Favorites", strUserProfile & "\Favorites_back"
   If Err Then
      objShell.Logevent 1, "Cannot backup local favorites" & vbCRLF & _
      "Error Description:" & Err.Description
      Err.Clear
      Wscript.Quit
   Else
      objFSO.DeleteFolder strUserProfile & "\Favorites"
      Err.Clear
      objFSO.CopyFolder strHomeShare & "\Favorites",strUserProfile & "\Favorites"
   End If
Else
   ObjShell.LogEvent 2, "Cannot locate server version of favorites" & vbCRLF & _
   "Error Description:" & strHomeShare & "\Facorites not found"
   Wscript.Quit
End If
   
'Record status to the event viewer
If Err Then
   objShell.LogEvent 1, "Cannot copy favorites from " & strHomeShare & _
   "\Favorites to " & strUserProfile & "\Favorites" & vbCRLF & _
   "Error Description: " & Err.Description
   Err.Clear
Else
   objShell.LogEvent 0, "Favorites successfully copied from " & strHomeShare & _
   "\Favorites to " & strUserProfile & "\Favorites"   
End If

'Close variables
Set objShell = Nothing
Set objFSO = Nothing