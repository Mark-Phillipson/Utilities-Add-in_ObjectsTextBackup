Option Compare Database
Option Explicit

'Constants for examining how a field is indexed.
Private Const intcIndexNone As Integer = 0
Private Const intcIndexGeneral As Integer = 1
Private Const intcIndexUnique As Integer = 3
Private Const intcIndexPrimary As Integer = 7

Type TableDetails
    TableName As String
    SourceTableName As String
    Attributes As Long
    IndexSQL As String
    Description As Variant
End Type

Sub FixConnectionsBroken(ServerName As String, DatabaseName As String, Optional strTable As String = "All", Optional booleanTrusted As Boolean = True _
, Optional stringUser As String = "", Optional stringPassword As String = "", _
Optional stringODBCDriverName As String = "ODBC Driver 13 for SQL Server", Optional blnAlwaysRemovePrefixCheckBox As Boolean = True, _
Optional stringPrefix As String = "dbo.")
' This code was originally written by
' Doug Steele, MVP  djsteele@canada.com
' You are free to use it in any application
' provided the copyright notice is left unchanged.
'
' Description:  This subroutine looks for any TableDef objects in the
'               database which have a connection string, and changes the
'               Connect property of those TableDef objects to use a
'               DSN-less connection.
'               This specific routine connects to the specified SQL Server
'               database on a specified server. It assumes trusted connection.
'
' Inputs:   ServerName:     Name of the SQL Server server (string)
'           DatabaseName:   Name of the database on that server (string)
' Sample call
' FixConnections "CARRIERAY\SQLEXPRESS", "SQLSE_AY"
' or
' FixConnections "CARRIERAY\SQLEXPRESS", "SQLSE_AY", "dbo_tblPCs"

On Error GoTo Err_FixConnections

Dim dbCurrent As DAO.database
Dim prpCurrent As DAO.Property
Dim tdfCurrent As DAO.tableDef
Dim intLoop As Integer
Dim intToChange As Integer
Dim strDescription As String
Dim typNewTables() As TableDetails

  intToChange = 0

  Set dbCurrent = CurrentDb ' DBEngine.Workspaces(0).Databases(0)

' Build a list of all of the connected TableDefs and
' the tables to which they're connected.
dbCurrent.TableDefs.Refresh
  For Each tdfCurrent In dbCurrent.TableDefs
    'Debug.Print tdfCurrent.Name
    If tdfCurrent.Name = strTable Or strTable = "All" Then
        If Len(tdfCurrent.Connect) > 0 And InStr(tdfCurrent.Connect, "mdb") = 0 Then
          ReDim Preserve typNewTables(0 To intToChange)
          typNewTables(intToChange).Attributes = tdfCurrent.Attributes
          typNewTables(intToChange).TableName = tdfCurrent.Name
          typNewTables(intToChange).SourceTableName = tdfCurrent.SourceTableName
          typNewTables(intToChange).IndexSQL = GenerateIndexSQL(tdfCurrent.Name)
          typNewTables(intToChange).Description = Null
          typNewTables(intToChange).Description = tdfCurrent.Properties("Description")
          intToChange = intToChange + 1
        End If
    End If
  Next
  

' Loop through all of the linked tables we found
  For intLoop = 0 To (intToChange - 1)

' Delete the existing TableDef object
    If typNewTables(intLoop).TableName = strTable Or strTable = "All" Then
        Dim stringOldName As String
        Dim stringOriginalName As String
        stringOldName = Right("_Old" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & typNewTables(intLoop).TableName, 64)
       Set tdfCurrent = dbCurrent.TableDefs(typNewTables(intLoop).TableName)
       stringOriginalName = tdfCurrent.Name
'       tdfCurrent.Name = stringOldName
      ' Set tdfCurrent = dbcurrent.CreateTableDef(typNewTables(intLoop).TableName)
        If booleanTrusted Then
            tdfCurrent.Connect = "ODBC;DRIVER={" & stringODBCDriverName & "};DATABASE=" & _
                                DatabaseName & ";SERVER=" & ServerName & _
                                ";Trusted_Connection=Yes;"
        Else
'            tdfCurrent.Connect = "ODBC;DRIVER={" & stringODBCDriverName & "};DATABASE=" & _
'                             DatabaseName & ";SERVER=" & ServerName & _
'                             ";Trusted_Connection=No;UID=" & stringUser & ";PWD=" & stringPassword & ";Connection Timeout=30;APP=Microsoft Office 2016;WSID=DESKTOP-H0A1CD1"
        
            tdfCurrent.Connect = "ODBC;DRIVER={" & stringODBCDriverName & "};" & _
              "DATABASE=" & DatabaseName & ";" & _
              "SERVER=" & ServerName & ";" & _
              "UID=" & stringUser & ";" & _
              "PWD=" & stringPassword & ";"
        End If
        If InStr(typNewTables(intLoop).SourceTableName, stringPrefix) = 0 Then
            'tdfCurrent.SourceTableName = stringPrefix & typNewTables(intLoop).SourceTableName
        Else
            'tdfCurrent.SourceTableName = typNewTables(intLoop).SourceTableName
        End If
        If Left(typNewTables(intLoop).SourceTableName, 4) = "" & Application.TempVars![CurrentDefaultSchema] & "." And blnAlwaysRemovePrefixCheckBox Then
            'tdfCurrent.Name = Mid(typNewTables(intLoop).SourceTableName, 5)
        End If
        On Error Resume Next
        'dbCurrent.TableDefs.Append tdfCurrent
        tdfCurrent.RefreshLink
        If Err.Number = 0 Then
            'dbcurrent.TableDefs.Delete stringOldName
        Else
            'tdfCurrent.Name = stringOriginalName
            'dbcurrent.TableDefs(stringOldName).Name = typNewTables(intLoop).TableName
            MsgBox "There was a problem when the system tried to link the following table " & typNewTables(intLoop).TableName _
             & Err.Description _
            , vbExclamation + vbOKOnly + vbDefaultButton1 _
            , "Problem Linking Table"
        End If
        
    
        ' Where it existed, add the Description property to the new table.
    
        'Put the server name into description
        strDescription = ServerName & " " & DatabaseName
        Set prpCurrent = tdfCurrent.CreateProperty("Description", dbText, strDescription)
        tdfCurrent.Properties.Append prpCurrent
    
        ' Where it existed, create the __UniqueIndex index on the new table.
    
        If Len(typNewTables(intLoop).IndexSQL) > 0 Then
      '    dbcurrent.Execute typNewTables(intLoop).IndexSQL, dbFailOnError
        End If
        Dim DAOField As DAO.field
        For Each DAOField In tdfCurrent.Fields
            If DAOField.Type = dbBoolean Then
                SetPropertyDAO DAOField, "DisplayControl", dbInteger, acCheckBox
            End If
        Next
        AddCommentsToLinkedTable tdfCurrent.Name, tdfCurrent.Connect, "SQL Server"
        AddCaptionsToLinkedTable tdfCurrent.Name, tdfCurrent.Connect
        


    End If
  Next
  Application.RefreshDatabaseWindow
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
    Case 3270
      Resume Next
    Case 3291
      MsgBox "Problem creating the Index using" & vbCrLf & _
          typNewTables(intLoop).IndexSQL, _
          vbOKOnly + vbCritical, "Fix Connections"
      Resume End_FixConnections
    Case Else
      MsgBox Err.Description & " (" & Err.Number & ") encountered", _
          vbOKOnly + vbCritical, "Fix Connections"
      Resume End_FixConnections
  End Select
Resume
End Sub



Sub ResetOraConnections(ServerName As String, DatabaseName As String, _
    strTable As String, strUser As String, strPwd As String)

Dim dbCurrent As DAO.database
Dim prpCurrent As DAO.Property
Dim tdfCurrent As DAO.tableDef
Dim intLoop As Integer
Dim intToChange As Integer
Dim strDescription As String
Dim typNewTables() As TableDetails

  intToChange = 0

  Set dbCurrent = DBEngine.Workspaces(0).Databases(0)

' Build a list of all of the connected TableDefs and
' the tables to which they're connected.
dbCurrent.TableDefs.Refresh
  'For Each tdfCurrent In dbCurrent.TableDefs
    Set tdfCurrent = dbCurrent.TableDefs(strTable)
    ''Debug.Print tdfCurrent.Name
    
    If Len(tdfCurrent.Connect) > 0 And InStr(tdfCurrent.Connect, "Oracle in") > 0 Then
      ReDim Preserve typNewTables(0 To intToChange)
      typNewTables(intToChange).Attributes = tdfCurrent.Attributes
      typNewTables(intToChange).TableName = tdfCurrent.Name
      typNewTables(intToChange).SourceTableName = tdfCurrent.SourceTableName
      typNewTables(intToChange).IndexSQL = GenerateIndexSQL(tdfCurrent.Name)
      typNewTables(intToChange).Description = Null
      'typNewTables(intToChange).Description = tdfCurrent.Properties("Description")
      intToChange = intToChange + 1
    End If
  'Next

' Loop through all of the linked tables we found

  For intLoop = 0 To (intToChange - 1)

' Delete the existing TableDef object

    If typNewTables(intLoop).TableName = strTable Then
    
        dbCurrent.TableDefs.Delete typNewTables(intLoop).TableName
    
        ' Create a new TableDef object, using the DSN-less connection
        
        '    Set tdfCurrent = dbCurrent.CreateTableDef(typNewTables(intLoop).TableName)
        '    tdfCurrent.Connect = "ODBC;Description=SQL Server 2005 Express;DRIVER=SQL Native Client;DATABASE=" & _
        '                        DatabaseName & ";SERVER=" & ServerName & _
        '                        ";Trusted_Connection=Yes;"
        
       'Set tdfCurrent = dbCurrent.CreateTableDef(typNewTables(intLoop).TableName)
       ' tdfCurrent.connect = "ODBC;DRIVER={sql server};DATABASE=" & _
        '                    DatabaseName & ";SERVER=" & ServerName & _
        '                    ";Trusted_Connection=Yes;"
        
        Set tdfCurrent = dbCurrent.CreateTableDef(typNewTables(intLoop).TableName)
        'tdfCurrent.connect = "ODBC;DRIVER={Oracle in ODACHome1};SERVER=XE;UID=bms;PWD=qwe;DBQ=XE;DBA=W;APA=T;EXC=F;XSM=Default;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=Me;CSR=F;FWC=F;FBS=60000;TLO=O;;TABLE=BMS.BMS_ACTIVE_JOB_DETAILS"
        tdfCurrent.Connect = "ODBC;DRIVER={Oracle in ODACHome1};" _
        & "SERVER=" & ServerName _
        & ";UID=" & strUser _
        & ";PWD=" & strPwd _
        & ";DBQ=" & DatabaseName _
        & ";TABLE=" & strTable
        
        tdfCurrent.SourceTableName = typNewTables(intLoop).SourceTableName
        dbCurrent.TableDefs.Append tdfCurrent
    
        ' Add the Description property to the new table. i.e. ServerName
    
        strDescription = ServerName
        Set prpCurrent = tdfCurrent.CreateProperty("Description", dbText, strDescription)
        tdfCurrent.Properties.Append prpCurrent
    
        ' Where it existed, create the __UniqueIndex index on the new table.
    
        If Len(typNewTables(intLoop).IndexSQL) > 0 Then
          dbCurrent.Execute typNewTables(intLoop).IndexSQL, dbFailOnError
        End If
    End If
  Next
    'ToggleHiddenProperty strTable, False
  
  Application.RefreshDatabaseWindow
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
    Case 3270
      Resume Next
    Case 3291
      MsgBox "Problem creating the Index using" & vbCrLf & _
          typNewTables(intLoop).IndexSQL, _
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
' Doug Steele, MVP  djsteele@canada.com
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
'db.Execute "CREATE INDEX " & sPrimaryKeyName & " ON " & sLocalTableName & "(" & sPrimaryKeyField & ") WITH PRIMARY;"


      If idxCurr.Fields.Count > 0 Then
        strSQL = "CREATE INDEX __UniqueIndex ON [" & TableName & "] ("
        For Each fldCurr In idxCurr.Fields
          strSQL = strSQL & "[" & fldCurr.Name & "], "
        Next

' Remove the trailing comma and space

        strSQL = Left$(strSQL, Len(strSQL) - 2) & ")"
        strSQL = strSQL & " WITH PRIMARY;"
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
Resume
End Function

Public Sub AddCommentsToLinkedTable(strTableName As String, stringConnection As String, Optional strType As String = "Oracle")
    Dim stringSQLText As String
    Dim strSQLtext  As String 'Aug 10
    Dim queryDefinition As DAO.QueryDef
    Dim db As DAO.database
    Dim codeDatabase As DAO.database
    Dim rs As DAO.Recordset
    Dim td As DAO.tableDef
    Dim fld As DAO.field
'    If Not strType = "Oracle" Then MsgBox "not set up for " & strType: Exit Sub
    On Error GoTo AddCommentsToLinkedTable_Error
    Set db = CurrentDb
    Set codeDatabase = codeDB
    Dim stringQueryName As String
    If InStr(stringConnection, ".database.windows.net") > 0 Then Exit Sub ' do not try adding descriptions and captions for Windows Azure
    If strType = "Oracle" Then
        Set queryDefinition = db.QueryDefs("qry_PTQ_User_Col_Comments")
        queryDefinition.Connect = stringConnection
        strSQLtext = "SELECT qry_PTQ_User_Col_Comments.TABLE_NAME" & vbCrLf
        strSQLtext = strSQLtext & "           , qry_PTQ_User_Col_Comments.COLUMN_NAME" & vbCrLf
        strSQLtext = strSQLtext & "           , qry_PTQ_User_Col_Comments.COMMENTS" & vbCrLf
        strSQLtext = strSQLtext & "        FROM qry_PTQ_User_Col_Comments" & vbCrLf
        strSQLtext = strSQLtext & "       WHERE (((qry_PTQ_User_Col_Comments.TABLE_NAME)='" & Mid(strTableName, 5) & "'))" & vbCrLf
        strSQLtext = strSQLtext & "    ORDER BY qry_PTQ_User_Col_Comments.COLUMN_NAME;"
    ElseIf strType = "SQL Server" Then
        stringQueryName = "qrySQLServerColumnDescriptions"
        Set queryDefinition = codeDatabase.QueryDefs(stringQueryName)
                                                                '   substring(cast( ep.value AS nvarchar(255)) ,0,255) AS
        stringSQLText = "select st.name [Table], sc.name [Column_Name],  Substring(cast(sep.value AS nvarchar(255)) ,0,255) AS [Comments]" _
         & " from sys.tables st    inner join sys.columns sc on st.object_id = sc.object_id    left join sys.extended_properties sep on st.object_id = sep.major_id and sc.column_id = sep.minor_id" _
         & " and sep.name = 'MS_Description'    where st.name ='" & strTableName & "'"
        queryDefinition.SQL = stringSQLText
        queryDefinition.Connect = stringConnection
'        strSQLtext = "SELECT qrySQLServerColumnDescriptions.ObjectName AS TABLE_NAME" & vbCrLf
'        strSQLtext = strSQLtext & "           , qrySQLServerColumnDescriptions.ColumnName AS COLUMN_NAME" & vbCrLf
'        strSQLtext = strSQLtext & "           , qrySQLServerColumnDescriptions.PropertyValue AS COMMENTS" & vbCrLf
'        strSQLtext = strSQLtext & "        FROM qrySQLServerColumnDescriptions" & vbCrLf
'        strSQLtext = strSQLtext & "       WHERE (((qrySQLServerColumnDescriptions.ObjectName)='" & strTableName & "'));"
        
    End If
    
    Set rs = queryDefinition.OpenRecordset()
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    Set td = db.TableDefs(strTableName)
    Do Until rs.EOF
        
        Set fld = td.Fields(rs![Column_Name])

        SetPropertyDAO fld, "Description", _
            dbText, rs![Comments]


        rs.MoveNext
    Loop
    rs.Close


ExitHere:
   Exit Sub

AddCommentsToLinkedTable_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure AddCommentsToLinkedTable of Module basLinkedTables" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Sub



Function CreateTableDAO()
    'Purpose:   Create two tables using DAO.
    Dim db As DAO.database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.field
    
    'Initialize the Contractor table.
    Set db = CurrentDb()
    Set tdf = db.CreateTableDef("tblDaoContractor")
    
    'Specify the fields.
    With tdf
        'AutoNumber: Long with the attribute set.
        Set fld = .CreateField("ContractorID", dbLong)
        fld.Attributes = dbAutoIncrField + dbFixedField
        .Fields.Append fld
        
        'Text field: maximum 30 characters, and required.
        Set fld = .CreateField("Surname", dbText, 30)
        fld.Required = True
        .Fields.Append fld
        
        'Text field: maximum 20 characters.
        .Fields.Append .CreateField("FirstName", dbText, 20)
        
        'Yes/No field.
        .Fields.Append .CreateField("Inactive", dbBoolean)
        
        'Currency field.
        .Fields.Append .CreateField("HourlyFee", dbCurrency)
        
        'Number field.
        .Fields.Append .CreateField("PenaltyRate", dbDouble)
        
        'Date/Time field with validation rule.
        Set fld = .CreateField("BirthDate", dbDate)
        fld.ValidationRule = "Is Null Or <=Date()"
        fld.ValidationText = "Birth date cannot be future."
        .Fields.Append fld
        
        'Memo field.
        .Fields.Append .CreateField("Notes", dbMemo)
        
        'Hyperlink field: memo with the attribute set.
        Set fld = .CreateField("Web", dbMemo)
        fld.Attributes = dbHyperlinkField + dbVariableField
        .Fields.Append fld
    End With
    
    'Save the Contractor table.
    db.TableDefs.Append tdf
    Set fld = Nothing
    Set tdf = Nothing
    Debug.Print "tblDaoContractor created."
    
    'Initialize the Booking table
    Set tdf = db.CreateTableDef("tblDaoBooking")
    With tdf
        'Autonumber
        Set fld = .CreateField("BookingID", dbLong)
        fld.Attributes = dbAutoIncrField + dbFixedField
        .Fields.Append fld
        
        'BookingDate
        .Fields.Append .CreateField("BookingDate", dbDate)
        
        'ContractorID
        .Fields.Append .CreateField("ContractorID", dbLong)
        
        'BookingFee
        .Fields.Append .CreateField("BookingFee", dbCurrency)
        
        'BookingNote: Required.
        Set fld = .CreateField("BookingNote", dbText, 255)
        fld.Required = True
        .Fields.Append fld
    End With
    
    'Save the Booking table.
    db.TableDefs.Append tdf
    Set fld = Nothing
    Set tdf = Nothing
    Debug.Print "tblDaoBooking created."
    
    'Clean up
    Application.RefreshDatabaseWindow   'Show the changes
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function

Function ModifyTableDAO()
    'Purpose:   How to add and delete fields to existing tables.
    'Note:      Requires the table created by CreateTableDAO() above.
    Dim db As DAO.database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.field
    
    'Initialize
    Set db = CurrentDb()

    Set tdf = db.TableDefs("tblDaoContractor")
    
    'Add a field to the table.
    tdf.Fields.Append tdf.CreateField("TestField", dbText, 80)
    Debug.Print "Field added."
    
    'Delete a field from the table.
    tdf.Fields.Delete "TestField"
    Debug.Print "Field deleted."
    
    'Clean up
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function

Function DeleteTableDAO()
    DBEngine(0)(0).TableDefs.Delete "DaoTest"
End Function

Function MakeGuidTable()
    'Purpose:   How to create a table with a GUID field.
    Dim db As DAO.database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.field
    Dim prp As DAO.Property

    Set db = CurrentDb()
    Set tdf = db.CreateTableDef("Table8")
    With tdf
        Set fld = .CreateField("ID", dbGUID)
        fld.Attributes = dbFixedField
        fld.DefaultValue = "GenGUID()"
        .Fields.Append fld
    End With
    db.TableDefs.Append tdf
End Function

Public Function CreateIndexDAO(tdf As tableDef, stringIndexName, stringFieldName, blnUnique As Boolean)
    Dim ind As DAO.index
    
    '1. Primary key index.
    On Error GoTo CreateIndexDAO_Error
    Set ind = tdf.CreateIndex(stringIndexName)
    With ind
        .Fields.Append .CreateField(stringFieldName)
        .Unique = blnUnique
       ' .Primary = True
    End With
    tdf.Indexes.Append ind
    
'    '2. Single-field index.
'    Set ind = tdf.CreateIndex("Inactive")
'    ind.Fields.Append ind.CreateField("Inactive")
'    tdf.Indexes.Append ind
'
'    '3. Multi-field index.
'    Set ind = tdf.CreateIndex("FullName")
'    With ind
'        .Fields.Append .CreateField("Surname")
'        .Fields.Append .CreateField("FirstName")
'    End With
'    tdf.Indexes.Append ind
    
    'Refresh the display of this collection.
    tdf.Indexes.Refresh
    
    'Clean up
    Set ind = Nothing

ExitHere:
   Exit Function

CreateIndexDAO_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure CreateIndexDAO of Module basLinkedTables" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function

Function DeleteIndexDAO()
    DBEngine(0)(0).TableDefs("tblDaoContractor").Indexes.Delete "Inactive"
End Function

Function CreateRelationDAO()
    Dim db As DAO.database
    Dim rel As DAO.relation
    Dim fld As DAO.field
    
    'Initialize
    Set db = CurrentDb()
    
    'Create a new relation.
    Set rel = db.CreateRelation("tblDaoContractortblDaoBooking")
    
    'Define its properties.
    With rel
        'Specify the primary table.
        .Table = "tblDaoContractor"
        'Specify the related table.
        .ForeignTable = "tblDaoBooking"
        'Specify attributes for cascading updates and deletes.
        .Attributes = dbRelationUpdateCascade + dbRelationDeleteCascade
        
        'Add the fields to the relation.
        'Field name in primary table.
        Set fld = .CreateField("ContractorID")
        'Field name in related table.
        fld.ForeignName = "ContractorID"
        'Append the field.
        .Fields.Append fld
        
        'Repeat for other fields if a multi-field relation.
    End With
    
    'Save the newly defined relation to the Relations collection.
    db.Relations.Append rel
    
    'Clean up
    Set fld = Nothing
    Set rel = Nothing
    Set db = Nothing
    Debug.Print "Relation created."
End Function

Function DeleteRelationDAO()
    DBEngine(0)(0).Relations.Delete "tblDaoContractortblDaoBooking"
End Function

Function DeleteQueryDAO()
    DBEngine(0)(0).QueryDefs.Delete "qryDaoBooking"
End Function

Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, _
    varValue As Variant, Optional strErrMsg As String) As Boolean
On Error GoTo ErrHandler
    'Purpose:   Set a property for an object, creating if necessary.
    'Arguments: obj = the object whose property should be set.
    '           strPropertyName = the name of the property to set.
    '           intType = the type of property (needed for creating)
    '           varValue = the value to set this property to.
    '           strErrMsg = string to append any error message to.
    
    If HasProperty(obj, strPropertyName) Then
        If Len(varValue & "") = 0 Then
            obj.Properties.Delete strPropertyName
        Else
            obj.Properties(strPropertyName) = varValue
        End If
    Else
        If (strPropertyName = "Format" Or strPropertyName = "Description") And Len(varValue & "") = 0 Then
            ' do nothing
        Else
            obj.Properties.Append obj.CreateProperty(strPropertyName, intType, Nz(varValue, ""))
        End If
    End If
    SetPropertyDAO = True

ExitHandler:
    Exit Function

ErrHandler:
    strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & " not set to " & varValue & _
        ". Error " & Err.Number & " - " & Err.Description & vbCrLf
        MsgBox strErrMsg
    Resume ExitHandler
    Resume
End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean
    'Purpose:   Return true if the object has the property.
    Dim varDummy As Variant
    
    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function

Function StandardProperties(strTableName As String)
    'Purpose:   Properties you always want set by default:
    '           TableDef:        Subdatasheets off.
    '           Numeric fields:  Remove Default Value.
    '           Currency fields: Format as currency.
    '           Yes/No fields:   Display as check box. Default to No.
    '           Text/memo/hyperlink: AllowZeroLength off,
    '                                UnicodeCompression on.
    '           All fields:      Add a caption if mixed case.
    'Argument:  Name of the table.
    'Note:      Requires: SetPropertyDAO()
    Dim db As DAO.database      'Current database.
    Dim tdf As DAO.tableDef     'Table nominated in argument.
    Dim fld As DAO.field        'Each field.
    Dim strCaption As String    'Field caption.
    Dim strErrMsg As String     'Responses and error messages.
    
    'Initalize.
    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTableName)
    
    'Set the table's SubdatasheetName.
    Call SetPropertyDAO(tdf, "SubdatasheetName", dbText, "[None]", _
        strErrMsg)
    
    For Each fld In tdf.Fields
        'Handle the defaults for the different field types.
        Select Case fld.Type
        Case dbText, dbMemo 'Includes hyperlinks.
            fld.AllowZeroLength = False
            Call SetPropertyDAO(fld, "UnicodeCompression", dbBoolean, _
                True, strErrMsg)
        Case dbCurrency
            fld.DefaultValue = 0
            Call SetPropertyDAO(fld, "Format", dbText, "Currency", _
                strErrMsg)
        Case dbLong, dbInteger, dbByte, dbDouble, dbSingle, dbDecimal
            fld.DefaultValue = vbNullString
        Case dbBoolean
            Call SetPropertyDAO(fld, "DisplayControl", dbInteger, _
                CInt(acCheckBox))
        End Select
        
        'Set a caption if needed.
        strCaption = ConvertMixedCase(fld.Name)
        If strCaption <> fld.Name Then
            Call SetPropertyDAO(fld, "Caption", dbText, strCaption)
        End If
        
        'Set the field's Description.
        Call SetFieldDescription(tdf, fld, , strErrMsg)
    Next
    
    'Clean up.
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    If Len(strErrMsg) > 0 Then
        Debug.Print strErrMsg
    Else
        Debug.Print "Properties set for table " & strTableName
    End If
End Function

Function ConvertMixedCase(ByVal strIn As String) As String
    'Purpose:   Convert mixed case name into a name with spaces.
    'Argument:  String to convert.
    'Return:    String converted by these rules:
    '           1. One space before an upper case letter.
    '           2. Replace underscores with spaces.
    '           3. No spaces between continuing upper case.
    'Example:   "FirstName" or "First_Name" => "First Name".
    Dim lngStart As Long        'Loop through string.
    Dim strOut As String        'Output string.
    Dim boolWasSpace As Boolean 'Last char. was a space.
    Dim boolWasUpper As Boolean 'Last char. was upper case.
    
    strIn = Trim$(strIn)        'Remove leading/trailing spaces.
    boolWasUpper = True         'Initialize for no first space.
    
    For lngStart = 1& To Len(strIn)
        Select Case Asc(Mid(strIn, lngStart, 1&))
        Case vbKeyA To vbKeyZ   'Upper case: insert a space.
            If boolWasSpace Or boolWasUpper Then
                strOut = strOut & Mid(strIn, lngStart, 1&)
            Else
                strOut = strOut & " " & Mid(strIn, lngStart, 1&)
            End If
            boolWasSpace = False
            boolWasUpper = True
            
        Case 95                 'Underscore: replace with space.
            If Not boolWasSpace Then
                strOut = strOut & " "
            End If
            boolWasSpace = True
            boolWasUpper = False
            
        Case vbKeySpace         'Space: output and set flag.
            If Not boolWasSpace Then
                strOut = strOut & " "
            End If
            boolWasSpace = True
            boolWasUpper = False
            
        Case Else               'Any other char: output.
            strOut = strOut & Mid(strIn, lngStart, 1&)
            boolWasSpace = False
            boolWasUpper = False
        End Select
    Next
    
    ConvertMixedCase = strOut
End Function

Function SetFieldDescription(tdf As DAO.tableDef, fld As DAO.field, _
Optional ByVal strDescrip As String, Optional strErrMsg As String) _
As Boolean
    'Purpose:   Assign a Description to a field.
    'Arguments: tdf = the TableDef the field belongs to.
    '           fld = the field to document.
    '           strDescrip = The description text you want.
    '                        If blank, uses Caption or Name of field.
    '           strErrMsg  = string to append any error messages to.
    'Notes:     Description includes field size, validation,
    '               whether required or unique.
    
    If (fld.Attributes And dbAutoIncrField) > 0& Then
        strDescrip = strDescrip & " Automatically generated " & _
            "unique identifier for this record."
    Else
        'If no description supplied, use the field's Caption or Name.
        If Len(strDescrip) = 0& Then
            If HasProperty(fld, "Caption") Then
                If Len(fld.Properties("Caption")) > 0& Then
                    strDescrip = fld.Properties("Caption") & "."
                End If
            End If
            If Len(strDescrip) = 0& Then
                strDescrip = fld.Name & "."
            End If
        End If
        
        'Size of the field.
        'Ignore Date, Memo, Yes/No, Currency, Decimal, GUID,
        '   Hyperlink, OLE Object.
'        Select Case fld.Type
'        Case dbByte, dbInteger, dbLong
'            strDescrip = strDescrip & " Whole number."
'        Case dbSingle, dbDouble
'            strDescrip = strDescrip & " Fractional number."
'        Case dbText
'            strDescrip = strDescrip & " " & fld.Size & "-char max."
'        End Select
        
        'Required and/or Unique?
        'Check for single-field index, and Required property.
'        Select Case IndexOnField(tdf, fld)
'        Case intcIndexPrimary
'            strDescrip = strDescrip & " Required. Unique."
'        Case intcIndexUnique
'            If fld.Required Then
'                strDescrip = strDescrip & " Required. Unique."
'            Else
'                strDescrip = strDescrip & " Unique."
'            End If
'        Case Else
'            If fld.Required Then
'                strDescrip = strDescrip & " Required."
'            End If
'        End Select
        
        'Validation?
'        If Len(fld.ValidationRule) > 0& Then
'            If Len(fld.ValidationText) > 0& Then
'                strDescrip = strDescrip & " " & fld.ValidationText
'            Else
'                strDescrip = strDescrip & " " & fld.ValidationRule
'            End If
'        End If
    End If
    
    If Len(strDescrip) > 0& Then
        strDescrip = Trim$(Left$(strDescrip, 255&))
        SetFieldDescription = SetPropertyDAO(fld, "Description", _
            dbText, strDescrip, strErrMsg)
    End If
End Function

Private Function IndexOnField(tdf As DAO.tableDef, fld As DAO.field) _
As Integer
    'Purpose:   Indicate if there is a single-field index _
    '               on this field in this table.
    'Return:    The constant indicating the strongest type.
    Dim ind As DAO.index
    Dim intReturn As Integer
    
    intReturn = intcIndexNone
    
    For Each ind In tdf.Indexes
        If ind.Fields.Count = 1 Then
            If ind.Fields(0).Name = fld.Name Then
                If ind.Primary Then
                    intReturn = (intReturn Or intcIndexPrimary)
                ElseIf ind.Unique Then
                    intReturn = (intReturn Or intcIndexUnique)
                Else
                    intReturn = (intReturn Or intcIndexGeneral)
                End If
            End If
        End If
    Next
    
    'Clean up
    Set ind = Nothing
    IndexOnField = intReturn
End Function

Function CreateQueryDAO()
    'Purpose:   How to create a query
    'Note:      Requires a table named MyTable.
    Dim db As DAO.database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb()
    
    'The next line creates and automatically appends the QueryDef.
    Set qdf = db.CreateQueryDef("qryMyTable")
    
    'Set the SQL property to a string representing a SQL statement.
    qdf.SQL = "SELECT MyTable.* FROM MyTable;"
    
    'Do not append: QueryDef is automatically appended!

    Set qdf = Nothing
    Set db = Nothing
    Debug.Print "qryMyTable created."
End Function

Function CreateDatabaseDAO()
    'Purpose:   How to create a new database and set key properties.
    Dim dbNew As DAO.database
    Dim prp As DAO.Property
    Dim strFile As String
    
    'Create the new database.
    strFile = "C:\SampleDAO.mdb"
    Set dbNew = DBEngine(0).CreateDatabase(strFile, dbLangGeneral)
    
    'Create example properties in new database.
    With dbNew
        Set prp = .CreateProperty("Perform Name AutoCorrect", dbLong, 0)
        .Properties.Append prp
        Set prp = .CreateProperty("Track Name AutoCorrect Info", _
            dbLong, 0)
        .Properties.Append prp
    End With
    
    'Clean up.
    dbNew.Close
    Set prp = Nothing
    Set dbNew = Nothing
    Debug.Print "Created " & strFile
End Function

Function ShowDatabaseProps()
    'Purpose:   List the properies of the current database.
    Dim db As DAO.database
    Dim prp As DAO.Property
    
    Set db = CurrentDb()
    For Each prp In db.Properties
        Debug.Print prp.Name
    Next
    
    Set db = Nothing
End Function

Function ShowFields(strTable As String)
    'Purpose:   How to read the fields of a table.
    'Usage:     Call ShowFields("Table1")
    Dim db As DAO.database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.field
    
    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTable)
    For Each fld In tdf.Fields
        Debug.Print fld.Name, FieldTypeName(fld)
    Next
    
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function

Function ShowFieldsRS(strTable)
    'Purpose:   How to read the field names and types from a table or query.
    'Usage:     Call ShowFieldsRS("Table1")
    Dim rs As DAO.Recordset
    Dim fld As DAO.field
    Dim strSQL As String
    
    strSQL = "SELECT " & strTable & ".* FROM " & strTable & " WHERE (False);"
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL)
    For Each fld In rs.Fields
        Debug.Print fld.Name, FieldTypeName(fld), "from " & fld.SourceTable & "." & fld.SourceField
    Next
    rs.Close
    Set rs = Nothing
End Function

Public Function FieldTypeName(fld As DAO.field)
    'Purpose: Converts the numeric results of DAO fieldtype to text.
    'Note:    fld.Type is Integer, but the constants are Long.
    Dim strReturn As String         'Name to return
    
    Select Case CLng(fld.Type)
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15
        
        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23
        
        'Constants for complex types don't work prior to Access 2007.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select
    
    FieldTypeName = strReturn
End Function

Function DAORecordsetExample()
    'Purpose:   How to open a recordset and loop through the records.
    'Note:      Requires a table named MyTable, with a field named MyField.
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT MyField FROM MyTable;"
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL)
    
    Do While Not rs.EOF
        Debug.Print rs!MyField
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function

Function ShowFormProperties(strFormName As String)
On Error GoTo Err_Handler
    'Purpose:   Loop through the controls on a form, showing names and properties.
    'Usage:     Call ShowFormProperties("Form1")
    Dim frm As Form
    Dim ctl As Control
    Dim prp As Property
    Dim strOut As String
    
    DoCmd.OpenForm strFormName, acDesign, WindowMode:=acHidden
    Set frm = Forms(strFormName)
    
    For Each ctl In frm
        For Each prp In ctl.Properties
            strOut = strFormName & "." & ctl.Name & "." & prp.Name & ": "
            strOut = strOut & prp.Type & vbTab
            strOut = strOut & prp.Value
            Debug.Print strOut
        Next
        If ctl.ControlType = acTextBox Then Stop
    Next
    
    Set frm = Nothing
    DoCmd.Close acForm, strFormName, acSaveNo

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 2186:
        strOut = strOut & Err.Description
        Resume Next
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ShowFormProperties()"
        Resume Exit_Handler
    End Select
End Function

Public Function ExecuteInTransaction(strSQL As String, Optional strConfirmMessage As String) As Long
On Error GoTo Err_Handler
    'Purpose:   Execute the SQL statement on the current database in a transaction.
    'Return:    RecordsAffected if zero or above.
    'Arguments: strSql = the SQL statement to be executed.
    '           strConfirmMessage = the message to show the user for confirmation. Number will be added to front.
    '           No confirmation if ZLS.
    '           -1 on error.
    '           -2 on user-cancel.
    Dim ws As DAO.Workspace
    Dim db As DAO.database
    Dim bInTrans As Boolean
    Dim bCancel As Boolean
    Dim strMsg As String
    Dim lngReturn As Long
    Const lngcUserCancel = -2&
    
    Set ws = DBEngine(0)
    ws.BeginTrans
    bInTrans = True
    Set db = ws(0)
    db.Execute strSQL, dbFailOnError
    lngReturn = db.RecordsAffected
    If strConfirmMessage <> vbNullString Then
        If MsgBox(lngReturn & " " & Trim$(strConfirmMessage), vbOKCancel + vbQuestion, "Confirm") <> vbOK Then
            bCancel = True
            lngReturn = lngcUserCancel
        End If
    End If
    
    'Commmit or rollback.
    If bCancel Then
        ws.Rollback
    Else
        ws.CommitTrans
    End If
    bInTrans = False

Exit_Handler:
    ExecuteInTransaction = lngReturn
    On Error Resume Next
    Set db = Nothing
    If bInTrans Then
        ws.Rollback
    End If
    Set ws = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ExecuteInTransaction()"
    lngReturn = -1
    Resume Exit_Handler
End Function

Function GetAutoNumDAO(strTable) As String
    'Purpose:   Get the name of the AutoNumber field, using DAO.
    Dim db As DAO.database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.field
    
    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTable)
    
    For Each fld In tdf.Fields
        If (fld.Attributes And dbAutoIncrField) <> 0 Then
            GetAutoNumDAO = fld.Name
            Exit For
        End If
    Next
    
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function

Public Function CreateFieldDAO(stringFieldName As String, stringTableName As String, IntegerSize As Integer, DataType As DataTypeEnum _
    , blnRequired As Boolean, blnAllowZeroLength As Boolean, stringCaption As String, stringDescription As String _
    , stringFormat As String, StringDefaultValue As String, integerDecimalPlaces As Integer, blnPlainText As Boolean, blnCreateUniqueIndex As Boolean) As Boolean
    Dim stringDatabase As String
    Dim blnAccessDatabase As Boolean
    Dim currentDatabase As DAO.database
    On Error GoTo CreateFieldDAO_Error
    Set currentDatabase = CurrentDb
    Dim tableDefinition As DAO.tableDef
    Dim externalDatabase As DAO.database
    Set tableDefinition = currentDatabase.TableDefs(stringTableName)
    If InStr(tableDefinition.Connect, "DATABASE=") > 0 And InStr(tableDefinition.Connect, "ODBC") = 0 Then
        blnAccessDatabase = True
        stringDatabase = Mid(tableDefinition.Connect, InStr(tableDefinition.Connect, "DATABASE=") + 9)
        Set externalDatabase = DBEngine(0).OpenDatabase(stringDatabase)
        Set tableDefinition = externalDatabase.TableDefs(tableDefinition.SourceTableName)
    Else
        blnAccessDatabase = False
        stringDatabase = ""
    End If
    Dim DAOField As DAO.field
    If DataType = dbDecimal Then
        CurrentProject.connection.Execute "ALTER TABLE [" & stringTableName & "] ADD COLUMN [" & stringFieldName & "] DECIMAL (" & IntegerSize & "," & integerDecimalPlaces & ");"
        currentDatabase.TableDefs.Refresh
        Set DAOField = tableDefinition.Fields(stringFieldName)
    Else
        Set DAOField = tableDefinition.CreateField(stringFieldName, DataType, IntegerSize)
    End If
    With DAOField
        .Required = blnRequired
        If DataType = dbText Then
            .AllowZeroLength = blnAllowZeroLength
        End If
        If Len(StringDefaultValue) > 0 Then
            .DefaultValue = IIf(Left(StringDefaultValue, 1) = "=", "", "=") & StringDefaultValue
        End If
    End With
    If Not DataType = dbDecimal Then
        tableDefinition.Fields.Append DAOField
    End If
        If Len(stringCaption & "") > 0 Then
            SetPropertyDAO DAOField, "Caption", dbText, stringCaption
        End If
        If Len(stringDescription & "") > 0 Then
            SetPropertyDAO DAOField, "Description", dbText, stringDescription
        End If
        If Len(stringFormat & "") > 0 Then
            SetPropertyDAO DAOField, "Format", dbText, stringFormat
        End If
        Select Case DataType
        Case dbCurrency, dbSingle, dbDouble
            If Len(integerDecimalPlaces & "") > 0 Then
                SetPropertyDAO DAOField, "DecimalPlaces", dbByte, integerDecimalPlaces
            End If
        Case dbBoolean
            SetPropertyDAO DAOField, "DisplayControl", dbInteger, acCheckBox
        Case dbMemo
            If blnPlainText Then
                SetPropertyDAO DAOField, "TextFormat", dbByte, acTextFormatPlain
            Else
                SetPropertyDAO DAOField, "TextFormat", dbByte, acTextFormatHTMLRichText
            End If
        End Select
    If blnCreateUniqueIndex Then
        CreateIndexDAO tableDefinition, DAOField.Name & "_Index", DAOField.Name, True
    End If
    CreateFieldDAO = True
ExitHere:
   Exit Function

CreateFieldDAO_Error:
    Select Case Err.Number
    Case 3211 ' The database engine could not locked the table because it is already in use
        MsgBox "The table appears to be open please close and retry.", vbExclamation, "Cannot Access Table"
        Resume ExitHere
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure CreateFieldDAO of Module basLinkedTables" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function

Public Sub DeleteFieldDAO(stringFieldName As String, stringTableName As String)
    Dim stringDatabase As String
    Dim blnAccessDatabase As Boolean
    Dim currentDatabase As DAO.database
    On Error GoTo DeleteFieldDAO_Error
    Set currentDatabase = CurrentDb
    Dim tableDefinition As DAO.tableDef
    Set tableDefinition = GetTableDefinition(currentDatabase, stringTableName)
    
    

    tableDefinition.Fields.Delete stringFieldName
    
    

ExitHere:
   Exit Sub

DeleteFieldDAO_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure DeleteFieldDAO of VBA Document Form_Browse Schema" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Sub


Public Sub RelinkSQLServerTable(stringTableName As String)
    Dim stringUsername As String
    Dim stringPasswordSQLServer As String
    Dim stringODBCDriverName As String
    Dim strSvrName As String
    Dim strSDBName As String
    ' this will relink a SQL Server linked table back to the same server and database for use when adding and removing columns
    Dim currentDatabase As DAO.database
    On Error GoTo RelinkSQLServerTable_Error
    Set currentDatabase = CurrentDb
    Dim tableDefinition As DAO.tableDef
    Set tableDefinition = currentDatabase.TableDefs(stringTableName)
    Dim stringODBCConnect As String
    stringODBCConnect = tableDefinition.Connect
    Dim stringName As String
    Dim stringSourceTableName As String
    stringSourceTableName = tableDefinition.SourceTableName
    Dim stringItems() As String
    Dim booleanTrustedConnection As Boolean
    booleanTrustedConnection = False
    stringItems = Split(stringODBCConnect, ";")
    Dim integerCounter As Integer
    For integerCounter = 0 To UBound(stringItems) - 1
        If InStr(stringItems(integerCounter), "DRIVER=") > 0 Then
            stringODBCDriverName = Mid(stringItems(integerCounter), InStr(stringItems(integerCounter), "DRIVER=") + 7)
        End If
        If InStr(stringItems(integerCounter), "SERVER=") > 0 Then
            strSvrName = Mid(stringItems(integerCounter), InStr(stringItems(integerCounter), "SERVER=") + 7)
        End If
        If InStr(stringItems(integerCounter), "DATABASE=") > 0 Then
            strSDBName = Mid(stringItems(integerCounter), InStr(stringItems(integerCounter), "DATABASE=") + 9)
        End If
        If InStr(stringItems(integerCounter), "Trusted_Connection=Yes") > 0 Then
            booleanTrustedConnection = True
        End If
        If InStr(stringItems(integerCounter), "UID=") > 0 Then
            stringUsername = Mid(stringItems(integerCounter), InStr(stringItems(integerCounter), "UID=") + 4)
        End If
        If InStr(stringItems(integerCounter), "PWD=") > 0 Then
            stringPasswordSQLServer = Mid(stringItems(integerCounter), InStr(stringItems(integerCounter), "PWD=") + 4)
        End If
        
    Next
    If Not booleanTrustedConnection Then
        If Len(stringUsername & "") = 0 Or Len(stringPasswordSQLServer) = 0 Then
            Dim stringFormName As String: stringFormName = "Login"
            If Not IsLoaded(stringFormName) Then
                DoCmd.Close acForm, stringFormName, acSaveYes
                DoCmd.OpenForm stringFormName, acNormal, , , acFormEdit, acDialog
            End If
            If Not IsLoaded(stringFormName) Then Exit Sub
            Dim AccessForm As Access.Form
            Set AccessForm = Forms(stringFormName)
            stringUsername = AccessForm![UsernameTextBox]
            stringPasswordSQLServer = AccessForm![PasswordTextBox]
        End If
    End If
    Dim booleanRemovePrefix As Boolean
    If Left(stringTableName, 3) = "dbo" Then
        booleanRemovePrefix = False
    Else
        booleanRemovePrefix = True
    End If
    SQLServer.FixConnections strSvrName, strSDBName, stringTableName, booleanTrustedConnection, Nz(stringUsername, ""), Nz(stringPasswordSQLServer, "") _
    , stringODBCDriverName, booleanRemovePrefix
    
    Set tableDefinition = currentDatabase.CreateTableDef(stringTableName)
    Dim DAOField As DAO.field
    For Each DAOField In tableDefinition.Fields
        If DAOField.Type = dbBoolean Then
            SetPropertyDAO DAOField, "DisplayControl", dbInteger, acCheckBox
        End If
    Next
ExitHere:
   Exit Sub

RelinkSQLServerTable_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure RelinkSQLServerTable" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Sub


Public Function GetSQLServerDataType(DataType As DataTypeEnum, IntegerSize As Integer, integerDecimalPlaces As Integer) As String
    Dim stringDataType As String
    Select Case DataType
    Case dbText
        stringDataType = "nvarchar(" & Trim(CStr(IntegerSize)) & ")"
    Case dbLong
        stringDataType = "int"
    Case dbMemo
        stringDataType = "nvarchar(max)"
    Case dbDate
        stringDataType = "datetime"
    Case dbBoolean
        stringDataType = "bit"
    Case dbCurrency
        stringDataType = "money"
    Case dbAttachment
        stringDataType = "varbinary(max)"
    Case dbSingle
        stringDataType = "float(24)"
    Case dbDouble
        stringDataType = "real"
    Case dbDecimal
        stringDataType = "decimal(" & IntegerSize & "," & integerDecimalPlaces & ")"
    End Select
    GetSQLServerDataType = stringDataType
End Function


Public Function GetApplicableDAODatabase(stringTableType As String, LocaltableDefinition As DAO.tableDef) As DAO.database
    Dim stringDatabase As String
        Select Case stringTableType
        Case "Access"
            Dim currentDatabase As DAO.database
            Set currentDatabase = CurrentDb
            Set GetApplicableDAODatabase = currentDatabase
        Case "Access Linked"
            Dim externalDatabase As DAO.database
            stringDatabase = Mid(LocaltableDefinition.Connect, InStr(LocaltableDefinition.Connect, "DATABASE=") + 9)
            Set externalDatabase = DBEngine(0).OpenDatabase(stringDatabase)
            Set GetApplicableDAODatabase = externalDatabase
        End Select

End Function

Public Sub AddCaptionsToLinkedTable(stringTableName As String, stringConnection As String)
    Dim currentDatabase As DAO.database
    Dim codeDatabase As DAO.database
    Dim Recordset As DAO.Recordset
    Dim stringSQLText As String
    On Error GoTo AddCaptionsToLinkedTable_Error
    If InStr(stringConnection, ".database.windows.net") > 0 Then Exit Sub ' do not try adding descriptions and captions for Windows Azure
    Set currentDatabase = CurrentDb
    Set codeDatabase = codeDB
    Dim queryDefinition As DAO.QueryDef
    Set queryDefinition = codeDatabase.QueryDefs("qrySQLserverColumnCaptions")
    queryDefinition.Connect = stringConnection
    stringSQLText = "SELECT qrySQLserverColumnCaptions.ObjectName AS ObjectName" & vbCrLf
    stringSQLText = stringSQLText & "           , qrySQLserverColumnCaptions.ColumnName" & vbCrLf
    stringSQLText = stringSQLText & "           , qrySQLserverColumnCaptions.PropertyName" & vbCrLf
    stringSQLText = stringSQLText & "           , qrySQLserverColumnCaptions.PropertyValue" & vbCrLf
    stringSQLText = stringSQLText & "        FROM qrySQLserverColumnCaptions" & vbCrLf
    stringSQLText = stringSQLText & "       WHERE (((qrySQLserverColumnCaptions.ObjectName)='" & stringTableName & "'));"
    Set Recordset = codeDatabase.OpenRecordset(stringSQLText, dbOpenSnapshot)
    If Not Recordset.EOF Then
        Recordset.MoveFirst
    End If
    Dim tableDefinition As DAO.tableDef
    Set tableDefinition = currentDatabase.TableDefs(stringTableName)
    Dim DAOField As DAO.field
    Do Until Recordset.EOF
        Set DAOField = tableDefinition.Fields(Recordset![ColumnName])
        SetPropertyDAO DAOField, "Caption", dbText, Recordset![PropertyValue]
        Recordset.MoveNext
    Loop
    Recordset.Close

ExitHere:
   Exit Sub

AddCaptionsToLinkedTable_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure AddCaptionsToLinkedTable of Module basLinkedTables" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

         
End Sub

Public Function ShowAccessBackEnd(stringLinkedTableName As String)
    Dim currentDatabase As DAO.database
    
    On Error GoTo ShowAccessBackEnd_Error
    Set currentDatabase = CurrentDb
    Dim tableDefinition As DAO.tableDef
    
    Set tableDefinition = currentDatabase.TableDefs(stringLinkedTableName)
'    Debug.Print tableDefinition.Connect
    Dim stringDatabase As String
    Dim stringResult As String
    If Len(tableDefinition.Connect & "") > 0 Then
        If InStr(tableDefinition.Connect, "ODBC") = 0 Then
            stringDatabase = Mid(tableDefinition.Connect, InStr(tableDefinition.Connect, "DATABASE=") + 9)
            stringResult = InputBox("Backend Database Name: " & stringDatabase, "Backend Database", stringDatabase)
        Else
            stringResult = InputBox("Connection String: " & tableDefinition.Connect, "ODBC Linked Table", tableDefinition.Connect)
        End If
    Else
        MsgBox "This table does not appear to be a linked table.", vbExclamation, "Linked Table"
    End If
    


ExitHere:
   Exit Function

ShowAccessBackEnd_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure ShowAccessBackEnd of Module basLinkedTables" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

    
End Function