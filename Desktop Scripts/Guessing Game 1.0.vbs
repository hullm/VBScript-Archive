'Created By Matthew Hull on some day a long time ago
'Documented on 4/25/04

'This is a fun script.  When I was first learning how to script I wanted to write a
'script that would incorporate many of the ideas I was reading about.  It is a simple
'guessing game.  Pick a number between one and ten.

'Force variable decliration
Option Explicit

'This constant determines the numeber of tries the user gets to guess
Const GAMEOVER = 5

Dim intRandomNumber,intGuess,intCount

'This subroutine generates a number between one and ten
Sub GenerateNumber()
   Randomize
   intRandomNumber = (Int(RND * 10)) + 1
   intRandomNumber = CInt(intRandomNumber)   
End Sub

'This subroutine gets a number from the user
Sub GetNumber()
   On Error Resume Next  
    
   Do
      
      'Ask the user to input a number
      intGuess = InputBox("Enter a number between 1 and 10","Guessing Game")
      
      'Change the input to an integer
      intGuess = CInt(intGuess)
      
      'If there is an error in the CInt command then the user didn't enter a number.
      'Then check to see if the number entered is between 1 and 10, if not display a 
      'message and change the number into a string
       If Err Then
         MsgBox "You did not enter a number, try again.",vbOkOnly,"Invalid Entry"
         Err.Clear
      ElseIf intGuess > 10 or intGuess < 1 Then
         MsgBox "You didn't enter a number between 1 and 10, try again",vbOkOnly,"Invalid Entry"
         intGuess = CStr(intGuess)
      End If
     
   'intGuess will only be an Integer if the user entered a proper value.  If they
   'didn't enter a proper value it will loop again    
   Loop Until VarType(intGuess) = vbInteger
End Sub

'Get a random number
GenerateNumber

'This will keep track of the number of tries it takes to guess
intCount = 0

'Display a message about this game
MsgBox "The object of this game is to pick a number between 1 in 10 in less then " & _
       GAMEOVER & " tries.",vbOkOnly,"Guessing Game"

'Loop until the user guesses the right numner or they have run out of tries
Do
   'Increase the number of tries by one
   intCount = intCount + 1
   
   'Get a number from the user
   GetNumber
   
   'Check to see if the user is out of tries if so exit the loop
   If intCount = GAMEOVER And intGuess <> intRandomNumber Then
      intCount = intCount + 1
      Exit Do
   End If
   
   'Display a message the informs the user weather their number was too high or too low
   If intGuess > intRandomNumber Then
      MsgBox "Incorrect, too high, try again. "& GameOver - intCount & " tries left.", _
      vbOkOnly,"Too High"
   ElseIf intGuess < intRandomNumber Then
      MsgBox "Incorrect, too low, try again. "& GameOver - intCount & " tries left.",_
      vbOkOnly,"Too Low"
   End If

Loop Until intGuess = intRandomNumber

'Display a message to the user based on weather or not they guessed the number
If intCount > GAMEOVER Then
   MsgBox "Forget it, you'll never get it.  Game Over.",vbOkOnly,"Game Over"
ElseIf intCount > 1 Then
   MsgBox "Congratulations!!! You Win. It took you " & intCount & " tries" ,vbOkOnly,"You Win"
Else
   MsgBox "Congratulations!!! You guessed it on the first try!!!",vbOkOnly,"You Win"
End If