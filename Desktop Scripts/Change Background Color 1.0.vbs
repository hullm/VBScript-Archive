'Created By Matthew Hull on some day a long time ago
'Documented on 4/25/04

'This script was created for a achool.  They wanted to change the background color for 
'the elem account whenever they changed the background picture.  This would allow them to enter the 
'RGB value of the color into an input box.  Then the script would write it into the regestry.

'Original Color 58 110 165

Dim WSHShell

'Create the Shell object, this will be used to write to the registry
Set WSHShell = WScript.CreateObject("Wscript.Shell")

'Create the core registry key
strRegKeyBGColor="HKEY_CURRENT_USER\Control Panel\Colors\Background"

'Get the current RGB color value
strCurrentValue = WSHShell.RegRead(strRegKeyBGColor)

'Get the new RGB color value from the user and display the old color in the box
strBGColor = InputBox("Enter the color in RGB format. IE 58 110 165","Enter RGB Color",strCurrentValue)

'Write the value to the registry
WSHShell.RegWrite strRegKeyBGColor, strBGColor