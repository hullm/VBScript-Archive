    ' *****************  BEGIN CODE HERE  ' *****************
    '
    Dim objScript
    Dim objAccess
    Dim strPathToMDB
    Dim strMsg

    ' ///////////// NOTE:  User must edit variables in this section /////
    '
    '  The following line of code is the only variable that need be edited
    '  You must provide a path to the Access MDB which will be compacted
    '
            strPathToMDB = "C:\servers.mdb"
    '
    ' ////////////////////////////////////////////////////////////////

    ' Set a name and path for a temporary mdb file
     strTempDB = "C:\Comp0001.mdb"

    ' Create Access 97 Application Object
    'Set objAccess = CreateObject("Access.Application.8")

    ' For Access 2000, use Application.9
    Set objAccess = CreateObject("Access.Application.11")

    ' Perform the DB Compact into the temp mdb file
    ' (If there is a problem, then the original mdb is  preserved)
    objAccess.DbEngine.CompactDatabase strPathToMDB ,strTempDB

    If Err.Number > 0 Then
        ' There was an error.  Inform the user and halt execution
        strMsg = "The following error was encountered while compacting database:"
        strMsg = strMsg & vbCrLf & vbCrLf & Err.Description
    Else
        ' Create File System Object to handle file manipulations
        Set objScript= CreateObject("Scripting.FileSystemObject")
    
        ' Back up the original file as Filename.mdbz.  In case of undetermined
        ' error, it can be recovered by simply removing the terminating "z".
        objScript.CopyFile strPathToMDB , strPathToMDB & "z", True

        ' Copy the compacted mdb by into the original file name
        objScript.CopyFile strTempDB, strPathToMDB, True

        ' We are finished with TempDB.  Kill it.
        objScript.DeleteFile strTempDB
    End If

    ' Always remember to clean up after yourself
    Set objAccess = Nothing
    Set objScript = Nothing
    '    
    ' ******************  END CODE HERE  ' ******************