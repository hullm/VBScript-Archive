'Created by Matthew Hull on 5/1/04

'This will generage a random password with lowercase letters and numbers

Option Explicit

On Error Resume Next

Dim intPasswordLength, intNumofPasswords, intNumofPasswordsLoop
Dim intPasswordLengthLoop, intRandomNumber, strPassword

Randomize Timer
'*************************************************************************
'Set the lengh and number of passwords you would like.

intPasswordLength = 8
intNumofPasswords = 50

'*************************************************************************

For intNumofPasswordsLoop = 1 to intNumofPasswords
   For intPasswordLengthLoop = 1 to intPasswordLength
      intRandomNumber = Int(2 * Rnd)   
      Select Case intRandomNumber
         Case 0 
            intRandomNumber = Int(10 * Rnd) + 48 '0-9
         Case 1
            intRandomNumber = Int(26 * Rnd) + 97 'a-z
      End Select   
      strPassword = strPassword & Chr(intRandomNumber)
   Next
   strPassword = strPassword & vbCRLF
Next
MsgBox strPassword,vbOkOnly,"Passwords"