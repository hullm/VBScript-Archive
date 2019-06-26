'Created by Matthew Hull on 10/11/03
'Documented on 4/24/03

'This will display network drives and there descriptions.  You can
'modify the descriptions to fit your network driver.

Option Explicit

'This sub will get the date in a DOW, Month Day, Year format where the Day
'of Week an month are written out in full format
Sub DateToday()

   Dim intWeekday, strWeekday, intMonth, strMonth, intDay, intYear

   'Get the Day of week and turn it into the full name
   intWeekday = Weekday(Date)
   strWeekday = WeekdayName(intWeekDay)
   
   'Get the month and turn it into the full name
   intMonth = Month(Date)
   strMonth = MonthName(intMonth)

   'Get the day of month and year
   intDay = Day(Date)
   intYear = Year(Date)
   
   'Build the date string
   strDate = strWeekday & ", " & strMonth & " " & intDay
   strDate = strDate & ", " & intYear
End Sub

'This constant make the code easier to read later
Const NETWORK_DRIVE = 3

'This group of constants contains the message you want displayed to the user
'if they have that drive mapped.
Const D_DRIVE = "There is no information available for this drive"
Const E_DRIVE = "There is no information available for this drive"
Const F_DRIVE = "There is no information available for this drive"
Const G_DRIVE = "There is no information available for this drive"
Const H_DRIVE = "There is no information available for this drive"
Const I_DRIVE = "There is no information available for this drive"
Const J_DRIVE = "There is no information available for this drive"
Const K_DRIVE = "There is no information available for this drive"
Const L_Drive = "There is no information available for this drive"
Const M_DRIVE = "There is no information available for this drive"
Const N_Drive = "There is no information available for this drive"
CONST O_DRIVE = "There is no information available for this drive"
Const P_DRIVE = "There is no information available for this drive"
Const Q_DRIVE = "There is no information available for this drive"
Const R_DRIVE = "There is no information available for this drive"
Const S_DRIVE = "There is no information available for this drive"
Const T_DRIVE = "There is no information available for this drive"
Const U_DRIVE = "There is no information available for this drive"
Const V_DRIVE = "There is no information available for this drive"
Const W_DRIVE = "There is no information available for this drive"
Const X_Drive = "There is no information available for this drive"
Const Y_DRIVE = "There is no information available for this drive"
Const Z_Drive = "There is no information available for this drive"
Const NO_INFO = "There is no information available for this drive"

Dim objFS, objDrives, objRE, strDriveMessage, Drive, strDate

'Create the File System Object which will be used to create the Drives Collection
Set objFS = CreateObject("Scripting.FileSystemObject")

'Create the Drives Collection, this is a collection of all the drives on the comptuer
Set objDrives = objFS.Drives

'Create a Regular Expression Object.  This will be used to check for network drives
Set objRE = New RegExp

'Create the start of the message that will display to the user
strDriveMessage = "The network drives on your computer are:" & vbCRLF & vbCRLF

'Loop thru each drive and check to see if it is a network drive
For each Drive in objDrives
   If Drive.DriveType = NETWORK_DRIVE Then
      strDriveMessage = strDriveMessage & Drive.Path & "\ " & vbTab
      
      'Check the drive letter and add the correct informaiton to the message
      Select Case Drive.Path
         Case "D:"
            strDriveMessage = strDriveMessage & D_DRIVE & vbCRLF
         Case "E:"
            strDriveMessage = strDriveMessage & E_DRIVE & vbCRLF
         Case "F:"
            strDriveMessage = strDriveMessage & F_DRIVE & vbCRLF
         Case "G:"
            strDriveMessage = strDriveMessage & G_DRIVE & vbCRLF
         Case "H:"
            strDriveMessage = strDriveMessage & H_DRIVE & vbCRLF
         Case "I:"
            strDriveMessage = strDriveMessage & I_DRIVE & vbCRLF
         Case "J:"
            strDriveMessage = strDriveMessage & J_DRIVE & vbCRLF
         Case "K:"
            strDriveMessage = strDriveMessage & K_DRIVE & vbCRLF
         Case "L:"
            strDriveMessage = strDriveMessage & L_DRIVE & vbCRLF
         Case "M:"
            strDriveMessage = strDriveMessage & M_DRIVE & vbCRLF
         Case "N:"
            strDriveMessage = strDriveMessage & N_DRIVE & vbCRLF
         Case "O:"
            strDriveMessage = strDriveMessage & O_DRIVE & vbCRLF
         Case "P:"
            strDriveMessage = strDriveMessage & P_DRIVE & vbCRLF
         Case "Q:"
            strDriveMessage = strDriveMessage & Q_DRIVE & vbCRLF
         Case "R:"
            strDriveMessage = strDriveMessage & R_DRIVE & vbCRLF
         Case "S:"
            strDriveMessage = strDriveMessage & S_DRIVE & vbCRLF
         Case "T:"
            strDriveMessage = strDriveMessage & T_DRIVE & vbCRLF
         Case "U:"
            strDriveMessage = strDriveMessage & U_DRIVE & vbCRLF
         Case "V:"
            strDriveMessage = strDriveMessage & V_DRIVE & vbCRLF
         Case "W:"
            strDriveMessage = strDriveMessage & W_DRIVE & vbCRLF
         Case "X:"
            strDriveMessage = strDriveMessage & X_DRIVE & vbCRLF
         Case "Y:"
            strDriveMessage = strDriveMessage & Y_DRIVE & vbCRLF
         Case "Z:"
            strDriveMessage = strDriveMessage & Z_DRIVE & vbCRLF
         Case Else
            strDriveMessage = strDriveMessage & No_Info & vbCRLF
         End Select
   End If
Next

'Set the pattern for the Reqular Expression
objRE.Pattern = "\\"

'Look for a "\" in the string, if there isn't one then there were no network
'drive found on the comptuer
If Not objRE.Test(strDriveMessage) Then
   strDriveMessage = "No Network Drives Found"
End If

'Go get the date and add it the message
DateToday
strDriveMessage = "Today is " & strDate & vbCRLF & strDriveMessage

'Display a message to the users
MsgBox strDriveMessage,vbOkOnly,"Network Drives"

'Colse the objects
Set objFS = Nothing
Set objDrives = Nothing
Set objRE = Nothing