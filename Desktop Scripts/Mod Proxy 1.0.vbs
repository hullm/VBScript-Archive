'Created by Matthew Hull a long time ago
'Documented on 5/23/04

'This script will write proxy settings to the registry.

Option Explicit

On Error Resume Next

Dim objWSHShell, strRegKeyRoot, strProxyServer, strProxyEnable

'Create the shell object this will be used to get OS information
Set objWSHShell = CreateObject("WScript.Shell")

'*****************************************************************************************
'Enter the proxy server name followed by a ":" then the port number

strProxyServer = "Server:8002"

'*****************************************************************************************

'This setting will be used to turn the proxy on
strProxyEnable = "1"

'This is the key that will hold the settings
strRegKeyRoot = "HKEY_CURRENT_USER\Software\Microsoft\"
strRegKeyRoot = strRegKeyRoot & "Windows\CurrentVersion\Internet Settings\"

'Write the values to the registry
WSHShell.RegWrite strRegKeyRoot & "ProxyServer", strProxyServer
WSHShell.RegWrite strRegKeyRoot & "ProxyEnable", strProxyEnable