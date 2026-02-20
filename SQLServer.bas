 Option Compare Database
Option Explicit
Type TableDetails
    TableName As String
    SourceTableName As String
    Attributes As Long
    IndexSQL As String
    Description As Variant
End Type


Public Function GetDefaultValueSQLServer(stringConnectionString As String, stringTableName As String, stringColumnName As String) As String
    Dim Command As ADODB.Command
    Dim longAffected As Long
    Dim stringSQLText As String
    On Error GoTo GetDefaultValueSQLServer_Error
    Set Command = New ADODB.Command
    Dim connection As ADODB.connection
    Set connection = New ADODB.connection
    connection.ConnectionString = stringConnectionString
    connection.Open
    Command.ActiveConnection = connection
    Command.CommandType = adCmdText
    stringSQLText = "SELECT SM.TEXT AS [DefaultValue] " & vbCrLf
    stringSQLText = stringSQLText & "        FROM " & Application.TempVars![CurrentDefaultSchema] & ".sysobjects SO " & vbCrLf
    stringSQLText = stringSQLText & "  INNER JOIN " & Application.TempVars![CurrentDefaultSchema] & ".syscolumns SC " & vbCrLf
    stringSQLText = stringSQLText & "          ON SO.id = SC.id " & vbCrLf
    stringSQLText = stringSQLText & "   LEFT JOIN " & Application.TempVars![CurrentDefaultSchema] & ".syscomments SM " & vbCrLf
    stringSQLText = stringSQLText & "          ON SC.cdefault = SM.id  " & vbCrLf
    stringSQLText = stringSQLText & "       WHERE SO.xtype = 'U'   And SO.NAME='" & stringTableName & "'  And SC.Name='" & stringColumnName & "'"
    Dim Recordset As New ADODB.Recordset
    Command.CommandText = stringSQLText
    Set Recordset = Command.Execute(longAffected)
    If Not Recordset.BOF And Not Recordset.EOF Then
        Recordset.MoveFirst
        GetDefaultValueSQLServer = Nz(Recordset![DefaultValue], "")
    End If
    Set Command = Nothing
    connection.Close
    Set connection = Nothing
ExitHere:
   Exit Function

GetDefaultValueSQLServer_Error:
    Select Case Err.Number
        Case -2147217865 ' invalid object name dbo.syscomments
        Resume ExitHere

    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure GetDefaultValueSQLServer of Module SQLServer" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

'***************** Code Start **************
Public Function TestFixedConnections()
    FixConnections "fgw4hyydwy.database.windows.net, 1433", "db_e811a279_8f27_472a_9bbd_af0dd58a6438", "db_e811a279_8f27_472a_9bbd_af0dd58a6438_ExternalWriter" _
    , "BZrTY0XZeJ6O9zD", "ODBC Driver 13 for SQL Server"
End Function
Sub FixConnections(ServerName As String, DatabaseName As String, Optional strTable As String = "All", Optional booleanTrusted As Boolean = True _
, Optional stringUser As String = "", Optional stringPassword As String = "", _
Optional stringODBCDriverName As String = "ODBC Driver 13 for SQL Server", Optional blnAlwaysRemovePrefixCheckBox As Boolean = True, _
Optional stringPrefix As String = "dbo.")
' This code was originally written by
' Doug Steele, MVP  AccessMVPHelp@gmail.com
' Modifications suggested by
' George Hepworth, MVP   ghepworth@gpcdata.com
'
' You are free to use it in any application
' provided the copyright notice is left unchanged.
'
' Description:  This subroutine looks for any TableDef objects in the
'               database which have a connection string, and changes the
'               Connect property of those TableDef objects to use a
'               DSN-less connection.
'               It then looks for any QueryDef objects in the database
'               which have a connection string, and changes the Connect
'               property of those pass-through queries to use the same
'               DSN-less connection.
'               This specific routine connects to the specified SQL Server
'               database on a specified server.
'               If a user ID and password are provided, it assumes
'               SQL Server Security is being used.
'               If no user ID and password are provided, it assumes
'               trusted connection (Windows Security).
'
' Inputs:   ServerName:     Name of the SQL Server server (string)
'           DatabaseName:   Name of the database on that server (string)
'           stringUser:            User ID if using SQL Server Security (string)
'           PWD:            Password if using SQL Server Security (string)
'
On Error GoTo Err_FixConnections

Dim dbCurrent As DAO.database
Dim prpCurrent As DAO.Property
Dim tdfCurrent As DAO.tableDef
Dim qdfCurrent As DAO.QueryDef
Dim intLoop As Integer
Dim intToChange As Integer
Dim strConnectionString As String
Dim strDescription As String
Dim strQdfConnect As String
Dim typNewTables() As TableDetails

' Start by checking whether using Trusted Connection or SQL Server Security

  If (Len(stringUser) > 0 And Len(stringPassword) = 0) Or (Len(stringUser) = 0 And Len(stringPassword) > 0) Then
    MsgBox "Must supply both User ID and Password to use SQL Server Security.", _
      vbCritical + vbOKOnly, "Security Information Incorrect."
    Exit Sub
  Else
    If Len(stringUser) > 0 And Len(stringPassword) > 0 And Not booleanTrusted Then

' Use SQL Server Security

      strConnectionString = "ODBC;DRIVER={" & stringODBCDriverName & "};" & _
        "DATABASE=" & DatabaseName & ";" & _
        "SERVER=" & ServerName & ";" & _
        "UID=" & stringUser & ";" & _
        "PWD=" & stringPassword & ";"
    Else

' Use Trusted Connection

      strConnectionString = "ODBC;DRIVER={" & stringODBCDriverName & "};" & _
        "DATABASE=" & DatabaseName & ";" & _
        "SERVER=" & ServerName & ";" & _
        "Trusted_Connection=YES;"
    End If
  End If

  intToChange = 0

  Set dbCurrent = DBEngine.Workspaces(0).Databases(0)

' Build a list of all of the connected TableDefs and
' the tables to which they're connected.
  dbCurrent.TableDefs.Refresh
  For Each tdfCurrent In dbCurrent.TableDefs
    If Len(tdfCurrent.Connect) > 0 Then
      If tdfCurrent.Name = strTable Or strTable = "All" Then
        ' sometimes we want to relink table from access to SQL Server (Boston Academic for example when using off-line)
        'If UCase$(Left$(tdfCurrent.Connect, 5)) = "ODBC;" Then
          ReDim Preserve typNewTables(0 To intToChange)
          typNewTables(intToChange).Attributes = tdfCurrent.Attributes
          typNewTables(intToChange).TableName = tdfCurrent.Name
          typNewTables(intToChange).SourceTableName = tdfCurrent.SourceTableName
          typNewTables(intToChange).IndexSQL = GenerateIndexSQL(tdfCurrent.Name)
          typNewTables(intToChange).Description = Null
          typNewTables(intToChange).Description = tdfCurrent.Properties("Description")
          intToChange = intToChange + 1
        'End If
      End If
    End If
  Next
' Warn if no tables found
    If intToChange = 0 Then
        MsgBox "Table Not Found in database: " & strTable & vbCrLf & "Connection String: " & strConnectionString, vbExclamation, "Table Not Found"
    End If
        

' Loop through all of the linked tables we found

  For intLoop = 0 To (intToChange - 1)
'Rename the existing table
    Dim tdfExisting As DAO.tableDef
    Set tdfExisting = dbCurrent.TableDefs(typNewTables(intLoop).TableName)
    
    tdfExisting.Name = typNewTables(intLoop).TableName & "~Old~"
' Create a new TableDef object, using the DSN-less connection

    Set tdfCurrent = dbCurrent.CreateTableDef(typNewTables(intLoop).TableName)
    tdfCurrent.Connect = strConnectionString

' Unfortunately, I'm current unable to test this code,
' but I've been told trying this line of code is failing for most people...
' If it doesn't work for you, just leave it out.
    On Error Resume Next
    tdfCurrent.Attributes = typNewTables(intLoop).Attributes Or DB_ATTACHSAVEPWD
    On Error GoTo Err_FixConnections
    ' include prefix if not already there i.e. dbo.
    If InStr(typNewTables(intLoop).SourceTableName, stringPrefix) = 0 Then
        tdfCurrent.SourceTableName = stringPrefix & typNewTables(intLoop).SourceTableName
    Else
        tdfCurrent.SourceTableName = typNewTables(intLoop).SourceTableName
    End If
    If Left(tdfCurrent.Name, 4) = "" & Application.TempVars![CurrentDefaultSchema] & "_" And blnAlwaysRemovePrefixCheckBox Then
        tdfCurrent.Name = Mid(tdfCurrent.Name, 5)
    End If

    dbCurrent.TableDefs.Append tdfCurrent

    ' Delete the existing TableDef object
    dbCurrent.TableDefs.Delete typNewTables(intLoop).TableName & "~Old~"


' Where it existed, add the Description property to the new table.

    If IsNull(typNewTables(intLoop).Description) = False Then
      strDescription = CStr(typNewTables(intLoop).Description)
      Set prpCurrent = tdfCurrent.CreateProperty("Description", dbText, strDescription)
      tdfCurrent.Properties.Append prpCurrent
    End If

' Where it existed, create the __UniqueIndex index on the new table.

    If Len(typNewTables(intLoop).IndexSQL) > 0 Then
      dbCurrent.Execute typNewTables(intLoop).IndexSQL, dbFailOnError
    End If
  Next
  
  
  
' Loop through all the QueryDef objects looked for pass-through queries to change.
' Note that, unlike TableDef objects, you do not have to delete and re-add the
' QueryDef objects: it's sufficient simply to change the Connect property.
' The reason for the changes to the error trapping are because of the scenario
' described in Addendum 6 below.

'  For Each qdfCurrent In dbCurrent.QueryDefs
'    On Error Resume Next
'    strQdfConnect = qdfCurrent.Connect
'    On Error GoTo Err_FixConnections
'    If Len(strQdfConnect) > 0 Then
'      If UCase$(Left$(qdfCurrent.Connect, 5)) = "ODBC;" Then
'        qdfCurrent.Connect = strConnectionString
'      End If
'    End If
'    strQdfConnect = vbNullString
'  Next qdfCurrent

End_FixConnections:
  Set tdfCurrent = Nothing
  Set dbCurrent = Nothing
  Exit Sub

Err_FixConnections:
' Specific error trapping added for Error 3291
' (Syntax error in CREATE INDEX statement.), since that's what many
' people were encountering with the old code.
' Also added error trapping for Error 3270 (Property Not Found.)
' to handle tables which don't have a description.

  Select Case Err.Number
    Case 3270, 3001 ' Invalid Arguments on attribute just ignore it
      Resume Next
    Case 3291
      MsgBox "Problem creating the Index using" & vbCrLf & _
        typNewTables(intLoop).IndexSQL, _
        vbOKOnly + vbCritical, "Fix Connections"
      Resume End_FixConnections
    Case 18456
      MsgBox "Wrong User ID or Password.", _
        vbOKOnly + vbCritical, "Fix Connections"
      Resume End_FixConnections
    Case Else
      MsgBox Err.Description & " (" & Err.Number & ") encountered", _
        vbOKOnly + vbCritical, "Fix Connections"
      Resume End_FixConnections
  End Select
Resume
End Sub

Function GenerateIndexSQL(TableName As String) As String
' This code was originally written by
' Doug Steele, MVP  AccessMVPHelp@gmail.com
' Modifications suggested by
' George Hepworth, MVP   ghepworth@gpcdata.com
'
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Description: Linked Tables should have an index __uniqueindex.
'              This function looks for that index in a given
'              table and creates an SQL statement which can
'              recreate that index.
'              (There appears to be no other way to do this!)
'              If no such index exists, the function returns an
'              empty string ("").
'
' Inputs:   TableDefObject: Reference to a Table (TableDef object)
'
' Returns:  An SQL string (or an empty string)
'

On Error GoTo Err_GenerateIndexSQL

Dim dbCurr As DAO.database
Dim idxCurr As DAO.index
Dim fldCurr As DAO.field
Dim strSQL As String
Dim tdfCurr As DAO.tableDef

  Set dbCurr = CurrentDb()
  Set tdfCurr = dbCurr.TableDefs(TableName)

  If tdfCurr.Indexes.Count > 0 Then

' Ensure that there's actually an index named
' "__UnigueIndex" in the table

    On Error Resume Next
    Set idxCurr = tdfCurr.Indexes("__uniqueindex")
    If Err.Number = 0 Then
      On Error GoTo Err_GenerateIndexSQL

' Loop through all of the fields in the index,
' adding them to the SQL statement

      If idxCurr.Fields.Count > 0 Then
        strSQL = "CREATE INDEX __UniqueIndex ON [" & TableName & "] ("
        For Each fldCurr In idxCurr.Fields
          strSQL = strSQL & "[" & fldCurr.Name & "], "
        Next

' Remove the trailing comma and space

        strSQL = Left$(strSQL, Len(strSQL) - 2) & ")"
      End If
    End If
  End If

End_GenerateIndexSQL:
  Set fldCurr = Nothing
  Set tdfCurr = Nothing
  Set dbCurr = Nothing
  GenerateIndexSQL = strSQL
  Exit Function

Err_GenerateIndexSQL:
' Error number 3265 is "Not found in this collection
' (in other words, either the tablename is invalid, or
' it doesn't have an index named __uniqueindex)
  If Err.Number <> 3265 Then
    MsgBox Err.Description & " (" & Err.Number & ") encountered", _
      vbOKOnly + vbCritical, "Generate Index SQL"
  End If
  Resume End_GenerateIndexSQL

End Function

'************** Code End *****************