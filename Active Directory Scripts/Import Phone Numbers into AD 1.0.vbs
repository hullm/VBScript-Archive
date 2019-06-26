'Created by Matthew Hull on 8/31/05

'This script will read a CSV file in the format lastname,firstname,phone number and try to
'locate the user in Active Directory and if found it will add the phone number.
'This script is lacking in error detection, use at your own risk.

Option Explicit

On Error Resume Next

Dim strSourceCSV, objRootDSE, objConnection, objCommand, objFSO, txtSourceCSV, strImportedData
Dim intCommaPosition, strLastName, strFirstName, strPhone, objRecordSet, objUser

'Enter the path the the source CSV
strSourceCSV = "C:\File.csv"

'Create a RootDSE object for the domain
Set objRootDSE = GetObject("LDAP://RootDSE")

'Establish a connection to Active Directory using ActiveX Data Object
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADSDSOObject;"

'Create the command object and attach it to the connection object
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection

'Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")


'Open a file and read each line
Set txtSourceCSV = objFSO.OpenTextFile(strSourceCSV)
While txtSourceCSV.AtEndOfLine = False
   strImportedData = txtSourceCSV.ReadLine

   'Then do something with strImportedData.  It will be in lastname,firstname,phone format.
   'Below is some code that will break it onto to variables.  It will need to be modified to do
   'three vars
   intCommaPosition = inStr(strImportedData,",")
   strLastName = Trim(Left(strImportedData,intCommaPosition - 1))
   strFirstName = Trim(Right(strImportedData,Len(strImportedData) - intCommaPosition))

   intCommaPosition = inStr(strFirstName,",")
   strPhone = Trim(Right(strFirstName,Len(strFirstName) - intCommaPosition))
   strFirstName = Trim(Left(strFirstName,intCommaPosition - 1))
   
   'Next we search AD for the user
   
   'objCommand defines the search base (in this case the whole domain) a filter
   '(All object of type user that match are user) and some attributes
   'associated with the returned objects (SamAccountName)
   objCommand.CommandText = "<LDAP://" & objRootDSE.Get("DefaultNamingContext") & _
   ">;(&(objectClass=user)(SN=" & strLastName & "));DistinguishedName,SN"
   
   'Initiate the LDAP query and return results to a RecordSet object.
   Set objRecordSet = objCommand.Execute
   
   'Loop thru each item in the Record Set and compare it to the last name
   'if they match then change the phone number.  Error detection will need
   'to be added.  If there is more then one user with the same last name.
   'Or you can do it amnually
   While Not objRecordset.EOF      
      If UCase(strLastName) = UCase(objRecordset.Fields("SN")) Then
         Set objUser = GetObject("LDAP://" & objRecordset.Fields("DistinguishedName"))
         If UCase(objUser.GivenName) = UCase(strFirstName) Then
            objUser.Put "telephonenumber", strPhone
            objUser.SetInfo
         End If
      End If
      objRecordset.MoveNext
   Wend
   'objRecordSet.MoveFirst      
Wend

objConnection.Close

MsgBox "Done"