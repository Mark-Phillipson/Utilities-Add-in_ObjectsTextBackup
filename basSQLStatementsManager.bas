Option Compare Database
Option Explicit

Public Sub DeleteQuery()
On Error GoTo DeleteQueryError
DoCmd.Close acQuery, "List Query", acSaveYes
DoCmd.DeleteObject A_QUERY, "List Query"
Exit Sub
DeleteQueryError:

DoCmd.Close acQuery, "List Query", acSaveNo
Resume Next

End Sub


Public Function CreateQueryFromSQLText(Optional strSQL As String)
    Dim db As DAO.database
    Dim Q As QueryDef
    'Dim ClipB As CClipboard
    On Error GoTo HandleErr
    Set db = CurrentDb
'    DoCmd.Close acQuery, "temp", acSaveYes
'    db.QueryDefs.Delete ("temp")
    
    
    'Set ClipB = New CClipboard
    If Len(strSQL & "") = 0 Then
        'DoCmd.RunCommand acCmdCopy
        strSQL = Nz(Screen.activeControl.Value, "")
    End If
    If Len(strSQL & "") = 0 Then
        MsgBox "There is no SQL selected to make a query.", vbOKOnly + vbExclamation + vbDefaultButton1, "Aborting"
        Exit Function
    End If
    Dim stringQueryName As String
    stringQueryName = "Temp" & Format(Now(), "yyyy-mmm-dd-hh-nn-ss")
    Set Q = db.CreateQueryDef(stringQueryName, strSQL)
    
    DoCmd.OpenQuery Q.Name, acViewDesign, acEdit
ExitHere:
    Exit Function

' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error
HandleErr:
    Select Case Err.Number
    Case 2046  'The command or action 'Copy' isn't available now.
        Resume ExitHere
    Case 3265 'Item not found in this collection.
        Resume Next
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.CreateQueryFromSQLText" 'ErrorHandler:$$N=basSQLStatementsManager.CreateQueryFromSQLText
        Resume ExitHere
    End Select
Resume
End Function

Public Function ShowImpSpec()
    Dim strImpSpec As String
    On Error GoTo HandleErr
    strImpSpec = Screen.activeControl
    
    DoCmd.OpenForm "frmImportSpecs", acNormal, , , acFormEdit, acWindowNormal, strImpSpec
ExitHere:
    Exit Function

' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error
HandleErr:
    Select Case Err.Number
    Case 2046 'Copy not available
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.ShowImpSpec"    'ErrorHandler:$$N=basSQLStatementsManager.ShowImpSpec
        Resume ExitHere
    End Select

End Function

Public Function OpenUKTable()
    If SysCmd(acSysCmdRuntime) Then
        DoCmd.OpenTable "tmpDefaultMailsortTable", acViewNormal, acReadOnly
    Else
        DoCmd.OpenTable "tmpDefaultMailsortTable", acViewNormal, acEdit
    End If
End Function

Public Function OpenOseasTable()
    If SysCmd(acSysCmdRuntime) Then
        DoCmd.OpenTable "tmpDefaultUnsortedTable", acViewNormal, acReadOnly
    Else
        DoCmd.OpenTable "tmpDefaultUnsortedTable", acViewNormal, acEdit
    End If
End Function

Public Function OpenProcessGDataPrep()
    DoCmd.OpenModule "basImportFiles", "ProcessGDataPrep"
End Function

Public Function OpenImportGeneral()
    DoCmd.OpenModule "basImportFiles", "ImportGeneral"
End Function

Public Function CaseConvert(Optional strType As String = "Toggle")
    Dim strConvert As String 'Added 16/06/2004
    Dim blnDatasheet As Boolean
    Dim ctl As Control
    Dim strCase As String ' changed from static to stop add-in from stopping access from closing as a test anyway08/10/2012
    On Error GoTo HandleErr
    blnDatasheet = False
    If Len(Screen.activeControl & "") > 0 Then
        strConvert = Screen.activeControl
        blnDatasheet = False
        Set ctl = Screen.activeControl
    Else
        If Len(Screen.ActiveDatasheet.activeControl & "") > 0 Then
            strConvert = Screen.ActiveDatasheet.activeControl
            blnDatasheet = True
            Set ctl = Screen.ActiveDatasheet.activeControl
            
        End If
    End If
    'If StrComp(strConvert, PCase(strConvert), vbBinaryCompare) = 0 Then strCase = "Proper Case"
    If StrComp(strConvert, UCase(strConvert), vbBinaryCompare) = 0 Then strCase = "UPPER"
    If StrComp(strConvert, LCase(strConvert), vbBinaryCompare) = 0 Then strCase = "lower"
    Select Case strCase
    Case ""
        strCase = "Proper Case"
    Case "Proper Case"
        strCase = "UPPER"
    Case "UPPER"
        strCase = "lower"
    Case "lower"
        strCase = "Proper Case"
    End Select
    If strType = "Toggle" Then strType = strCase
    
    
    Select Case strType
    Case "Proper Case"
        strConvert = "" 'PCase(strConvert)
    Case "UPPER"
        strConvert = UCase(strConvert)
    Case "lower"
        strConvert = LCase(strConvert)
    End Select
    ctl.Value = strConvert
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484 ' There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.CreateQueryFromSQLText" 'ErrorHandler:$$N=basSQLStatementsManager.CreateQueryFromSQLText
        Resume ExitHere
    End Select
Resume

End Function

Public Function AnalysColumn()
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim QD As DAO.QueryDef
    Dim strTable As String
    Dim strColumn As String
    
    On Error GoTo HandleErr
    strTable = Screen.ActiveDatasheet.Name
    
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    strSQLtext = "SELECT [" & strTable & "].[" & strColumn & "], Count([" & strTable & "].[" & strColumn & "]) AS Records " & vbCrLf
    strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
    strSQLtext = strSQLtext & "    GROUP BY [" & strTable & "].[" & strColumn & "];"
    Set db = CurrentDb
    DoCmd.Close acQuery, "temp", acSaveYes
    On Error Resume Next
    DoCmd.DeleteObject acQuery, "temp"
    On Error GoTo HandleErr
    Set QD = db.CreateQueryDef("temp", strSQLtext)
    DoCmd.OpenQuery "temp", acViewDesign, acEdit

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select
Resume ' debug
End Function

Public Function LookUpATable(strTable As String, strColumn As String, _
    strOperator As String, Optional strSearch1 As String)
    Dim strWhere As String 'Added 19/09/2005
    Dim strSearch As String
    On Error GoTo HandleErr
    
    DoCmd.RunCommand acCmdSaveRecord
    If Len(strSearch1) = 0 Then
        strSearch = Nz(Screen.ActiveDatasheet.activeControl.Value, "")
    Else
        strSearch = strSearch1
    End If
    If strOperator = "Like" Then
        strWhere = "[" & strColumn & "] Like '*" & strSearch & "*'"
    Else
        strWhere = "[" & strColumn & "] = '" & strSearch & "'"
    End If
    DoCmd.OpenForm "frmCity-Country-Lookup", acNormal, , strWhere, acFormEdit, acDialog, strSearch
    

ExitHere:
    Exit Function

' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error
HandleErr:
    Select Case Err.Number
    Case 2046 'The command or action 'SaveRecord' isn't available now.
        Resume Next
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select

Resume
End Function

Public Function CopyFromTempQuery()
    Dim db As DAO.database
    Dim Q As DAO.QueryDef
On Error GoTo HandleErr
    
    Set db = CurrentDb
    Set Q = db.QueryDefs("temp")
    Screen.activeControl.Value = Q.SQL

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 3265 'Item not found in this collection.
        MsgBox "No Temp Query found.", vbOKOnly + vbExclamation + vbDefaultButton1, "Aborting..."
        Resume ExitHere

    Case Else
            MsgBox "Unexpected Error " & Err.Number & ": " & Err.Description, vbCritical, "basSQLStatementsManager.CopyFromTempQuery"   'ErrorHandler:$$N=basSQLStatementsManager.CopyFromTempQuery
        Resume ExitHere
    End Select
    Resume 'Debug Only
End Function

Public Function OpenZoomBox()

On Error GoTo HandleErr
    DoCmd.RunCommand acCmdZoomBox
    
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case '
    Case Else
            MsgBox "Unexpected Error " & Err.Number & ": " & Err.Description, vbCritical, "basSQLStatementsManager.OpenZoomBox" 'ErrorHandler:$$N=basSQLStatementsManager.OpenZoomBox
        Resume ExitHere
    End Select
    Resume 'Debug Only
End Function

Public Function BreakIntoNewTables()
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim strTable As String
    Dim strColumn As String
    Dim strMsg As String, k As Integer
    Dim rs As DAO.Recordset
    Dim intLongest As Integer
    Dim strNewTable As String
    Dim blnValidName As Boolean, lngRecords As Long
    Dim strCrap As String, lngNewTables As Long
    On Error GoTo HandleErr
    Set db = CurrentDb
    strTable = Screen.ActiveDatasheet.Name
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    db.TableDefs.Refresh
    If db.TableDefs(strTable).Fields(strColumn).Type <> dbText Then Exit Function
    strSQLtext = "SELECT Max(Len([" & strTable & "]![" & strColumn & "])) AS Longest" & vbCrLf
    strSQLtext = strSQLtext & "        FROM [" & strTable & "];"
    Set rs = db.OpenRecordset(strSQLtext)
    If rs.EOF Or IsNull(rs![Longest]) Then Exit Function
    intLongest = Len(strTable) + rs![Longest]
    rs.Close
    If intLongest > 64 Then
        MsgBox "Please note the table name and longest value are longer than the maximum length for a table name." _
        , vbExclamation, "Aborting..."
        Exit Function
    End If
    strSQLtext = "SELECT IIf(Len([" & strTable & "]![" & strColumn & "] & """")>0,[" & strTable & "]![" & strColumn & "],""Blank"") AS Split" & vbCrLf
    strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
    strSQLtext = strSQLtext & "    GROUP BY IIf(Len([" & strTable & "]![" & strColumn & "] & """")>0,[" & strTable & "]![" & strColumn & "],""Blank"");"
    Set rs = db.OpenRecordset(strSQLtext)
    rs.MoveLast
    If rs.RecordCount = 1 Or rs.RecordCount > 20 Then
            MsgBox "There is only 1 or more than 20 tables that would be created." _
        , vbExclamation, "Aborting... Table(s) " & rs.RecordCount
        Exit Function
    End If
    lngNewTables = rs.RecordCount
    rs.MoveFirst
    blnValidName = True
    Do Until rs.EOF
        For k = 1 To Len(rs![Split])
            Select Case Mid(rs![Split], k, 1)
            Case ".", ",", "/", "\", "*", """", "'", "!", "`", "[", "]"
                blnValidName = False
                strCrap = strCrap & rs![Split] & " Character " & Mid(rs![Split], k, 1) & vbCrLf
            End Select
        Next
        rs.MoveNext
    Loop
    If Not blnValidName Then
        MsgBox "There are characters in the split column that are not allowed in table names:" _
             & vbCrLf & vbCrLf & strCrap _
             & vbCrLf & vbCrLf & "Please removed and try again." _
        , vbOKOnly + vbExclamation + vbDefaultButton1, "Aborting..."
        Exit Function
    End If
    rs.MoveFirst
    Do Until rs.EOF
        'Build new table
        strNewTable = strTable & " " & rs![Split]
        If rs![Split] = "Blank" Then
            strSQLtext = "SELECT [" & strTable & "].* INTO [" & strNewTable & "]" & vbCrLf
            strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
            strSQLtext = strSQLtext & "       WHERE (((Len([" & strTable & "]![" & strColumn & "] & """"))=0));"
        Else
            strSQLtext = "SELECT [" & strTable & "].* INTO [" & strNewTable & "]" & vbCrLf
            strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
            strSQLtext = strSQLtext & "       WHERE ((([" & strTable & "].[" & strColumn & "])='" & rs![Split] & "'));"
        End If
        db.Execute strSQLtext
        strMsg = strMsg & "New Table '" & strNewTable & "' " & db.RecordsAffected & " rows" & vbCrLf
        lngRecords = lngRecords + db.RecordsAffected
        rs.MoveNext
    Loop
    RefreshDatabaseWindow
    MsgBox strMsg & vbCrLf & "Tables " & lngNewTables & " Rows " & lngRecords, vbInformation, "Break into New Tables Done..."

ExitHere:
    On Error Resume Next
    rs.Close
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case 3010 ' Table 'xxx' Already Exists.
        Select Case MsgBox("Table Already Exists." _
             & vbCrLf & "" _
             & vbCrLf & "Do you want to overwite " & strNewTable & "?" _
        , vbYesNo + vbExclamation + vbDefaultButton2, "Table Exists")
        Case vbYes
            DoCmd.DeleteObject acTable, strNewTable
            Resume
        Case vbNo
            Resume ExitHere
        End Select
    Case 3265 'Item not found in this collection.
        'looks like we are no in a table
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.BreakIntoNewTables"
        Resume ExitHere
    End Select
Resume ' debug
End Function

Public Function SubColumn()
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim QD As DAO.QueryDef
    Dim strTable As String
    Dim strColumn As String
    Dim rs As DAO.Recordset
    Dim fld As DAO.field
    On Error GoTo HandleErr
    strTable = Screen.ActiveDatasheet.Name
    
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strTable, dbOpenSnapshot)
    Set fld = rs.Fields(strColumn)
    Select Case fld.Type
    Case dbInteger, dbCurrency, dbLong, dbDecimal, dbDouble, dbSingle
        strSQLtext = "SELECT Sum([" & strTable & "].[" & strColumn & "]) AS [Sum Of " & strColumn & "]" & vbCrLf
        strSQLtext = strSQLtext & "        FROM [" & strTable & "];" & vbCrLf
        
        DoCmd.Close acQuery, "temp", acSaveYes
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "temp"
        On Error GoTo HandleErr
        Set QD = db.CreateQueryDef("temp", strSQLtext)
        DoCmd.OpenQuery "temp", acViewNormal, acEdit
    Case Else
        ShowBalloonTooltip "Error - Cannot Create Query", "Please check that the " _
        & strColumn & " column is one of the NUMBER data types!", btCritical
    End Select
    

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select
Resume ' debug
End Function

Public Function RunSQLTextShowErrors(Optional strSQL As String)
    Dim db As DAO.database
    Dim Q As QueryDef
    'Dim ClipB As CClipboard
    On Error GoTo HandleErr
    Set db = CurrentDb
    DoCmd.Close acQuery, "temp7", acSaveYes
    
    db.QueryDefs.Delete ("temp7")
    
    
    'Set ClipB = New CClipboard
    If Len(strSQL & "") = 0 Then
        'DoCmd.RunCommand acCmdCopy
        strSQL = Nz(Screen.activeControl.Value, "")
    End If
    If Len(strSQL & "") = 0 Then
        MsgBox "There is no SQL selected to make a query.", vbOKOnly + vbExclamation + vbDefaultButton1, "Aborting"
        Exit Function
    End If
    Set Q = db.CreateQueryDef("temp7", strSQL)
    Q.Execute dbFailOnError + dbSeeChanges
    MsgBox "Records Affected " & Q.RecordsAffected, vbInformation, "Query Has Run"
ExitHere:
    Exit Function

' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error
HandleErr:
    Select Case Err.Number
    Case 2046  'The command or action 'Copy' isn't available now.
        Resume ExitHere
    Case 3265 'Item not found in this collection.
        Resume Next
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description & Errors(0).Description, vbCritical, "Unexpected Error in basSQLStatementsManager.CreateQueryFromSQLText" 'ErrorHandler:$$N=basSQLStatementsManager.CreateQueryFromSQLText
        Resume ExitHere
    End Select
Resume
End Function

Public Function CreateUpdateQueryOnColumn()
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim QD As DAO.QueryDef
    Dim strTable As String
    Dim strColumn As String
    
    On Error GoTo HandleErr
    strTable = Screen.ActiveDatasheet.Name
    
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    
    strSQLtext = "UPDATE [" & strTable & "] SET [" & strTable & "].[" & strColumn & "] = '" & Screen.ActiveDatasheet.activeControl.Value & "';"

    Set db = CurrentDb
    'DoCmd.Close acQuery, "temp", acSaveYes
    'On Error Resume Next
    'DoCmd.DeleteObject acQuery, "temp"
    On Error GoTo HandleErr
    Dim stringQueryName As String
    stringQueryName = "TempUpdate" & Format(Now(), "yyyy-mmm-dd-hh-nn-ss")
    Set QD = db.CreateQueryDef(stringQueryName, strSQLtext)
    DoCmd.OpenQuery stringQueryName, acViewDesign, acEdit

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select
Resume ' debug
End Function

Public Function CreateAppendQueryOnColumn()
    Dim stringSQLText As String
    Dim db As DAO.database
    Dim QD As DAO.QueryDef
    Dim strTable As String
    Dim strColumn As String
    
    On Error GoTo HandleErr
    strTable = Screen.ActiveDatasheet.Name
    
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    stringSQLText = "INSERT INTO [" & strTable & "] ( [" & strTable & "].[" & strColumn & "] ) SELECT [" & strTable & "].[" & strColumn & "]" & vbCrLf
    stringSQLText = stringSQLText & "        FROM [" & strTable & "];"
    
    Set db = CurrentDb
'    On Error Resume Next
'    DoCmd.DeleteObject acQuery, "temp"
'    On Error GoTo HandleErr
    Dim stringQueryName As String
    stringQueryName = "TempAppend" & Format(Now(), "yyyy-mmm-dd-hh-nn-ss")
    Set QD = db.CreateQueryDef(stringQueryName, stringSQLText)
    DoCmd.OpenQuery stringQueryName, acViewDesign, acEdit

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select
Resume ' debug
End Function

Public Function CreateSelectQueryOnColumn()
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim QD As DAO.QueryDef
    Dim strTable As String
    Dim strColumn As String
    
    On Error GoTo HandleErr
    strTable = Screen.ActiveDatasheet.Name
    
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    
    
    strSQLtext = "SELECT [" & strTable & "].[" & strColumn & "]" & vbCrLf
    strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
    strSQLtext = strSQLtext & "       WHERE ((([" & strTable & "].[" & strColumn & "])='" & Screen.ActiveDatasheet.activeControl.Value & "'));"
    
    Set db = CurrentDb
'    DoCmd.Close acQuery, "temp", acSaveYes
'    On Error Resume Next
'    DoCmd.DeleteObject acQuery, "temp"
'    DoCmd.DeleteObject acQuery, "temp"
    'this should delete the Query in both the add-in and the current database
    On Error GoTo HandleErr
    Dim stringQueryName As String
    stringQueryName = "TempSelect" & Format(Now(), "yyyy-mmm-dd-hh-nn-ss")
    Set QD = db.CreateQueryDef(stringQueryName, strSQLtext)
    DoCmd.OpenQuery stringQueryName, acViewDesign, acEdit

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select
Resume ' debug
End Function


Public Function CreateDeleteQueryOnColumn()
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim QD As DAO.QueryDef
    Dim strTable As String
    Dim strColumn As String
    
    On Error GoTo HandleErr
    strTable = Screen.ActiveDatasheet.Name
    
    strColumn = Screen.ActiveDatasheet.activeControl.Name
    
    
    strSQLtext = "DELETE [" & strTable & "].[" & strColumn & "]" & vbCrLf
    strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
    strSQLtext = strSQLtext & "       WHERE ((([" & strTable & "].[" & strColumn & "])='" & Screen.ActiveDatasheet.activeControl.Value & "'));"
    
    Set db = CurrentDb
'    DoCmd.Close acQuery, "temp", acSaveYes
'    On Error Resume Next
'    DoCmd.DeleteObject acQuery, "temp"
'    On Error GoTo HandleErr
    Dim stringQueryName As String
    stringQueryName = "TempDelete" & Format(Now(), "yyyy-mmm-dd-hh-nn-ss")
    Set QD = db.CreateQueryDef(stringQueryName, strSQLtext)
    DoCmd.OpenQuery stringQueryName, acViewDesign, acEdit

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basSQLStatementsManager.AnalysColumn"   'ErrorHandler:$$N=basSQLStatementsManager.AnalysColumn
        Resume ExitHere
    End Select
Resume ' debug
End Function


Public Function SaveSQLtoTempQuery(Optional stringSQLText As String = "")
    Dim blnGetFormClipboard As Boolean
    On Error Resume Next
    If Len(stringSQLText & "") = 0 Then
        stringSQLText = Screen.activeControl.Value
    End If
    If Len(stringSQLText & "") = 0 Then
        stringSQLText = Screen.ActiveDatasheet.activeControl.Value
    End If
    If Len(stringSQLText & "") = 0 Then
        stringSQLText = Screen.ActiveForm.activeControl.Value
    End If
    On Error GoTo SaveSQLtoTempQuery_Error
    If Len(stringSQLText & "") = 0 Then
        blnGetFormClipboard = True
    End If
    Dim currentDatabase As DAO.database
    Set currentDatabase = CurrentDb
    DoCmd.Close acQuery, "temp", acSaveYes
    If Not blnGetFormClipboard Then
        On Error Resume Next
        DoCmd.DeleteObject acQuery, "temp"
        DoCmd.DeleteObject acQuery, "temp"
        On Error GoTo SaveSQLtoTempQuery_Error
        Dim queryDefinition As DAO.QueryDef
        Set queryDefinition = currentDatabase.CreateQueryDef("Temp", stringSQLText)
        DoCmd.OpenQuery "Temp", acViewDesign, acEdit
    End If
    
    If blnGetFormClipboard Then
        DoCmd.OpenQuery "Temp", acViewDesign, acEdit
        SendKeys "%HWQ", True ' switch to SQL view for the query
'        Dim longCounter As Long
'        For longCounter = 1 To 60000000
'            'Dummy
'        Next
        'DoCmd.RunCommand acCmdPaste
    End If

ExitHere:
   Exit Function

SaveSQLtoTempQuery_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SaveSQLtoTempQuery of Module basSQLStatementsManager" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only

End Function

Public Function LoadSQLfromTemporaryQuery()

    On Error GoTo LoadSQLfromTemporaryQuery_Error
    Dim stringSQLText As String
    DoCmd.Close acQuery, "temp", acSaveYes
    Dim currentDatabase As DAO.database
    Set currentDatabase = CurrentDb
    Dim queryDefinition As DAO.QueryDef
    Set queryDefinition = currentDatabase.QueryDefs("Temp")
    Screen.activeControl.Value = queryDefinition.SQL
    On Error GoTo LoadSQLfromTemporaryQuery_Error
ExitHere:
   Exit Function

LoadSQLfromTemporaryQuery_Error:
    Select Case Err.Number
    Case 3163 ' The field is too small to set the amount of data you are tempted to at try inserting or pasting less data
        MsgBox "The field is too small to accept the SQL", vbExclamation, "Data will not Fit"
         Resume ExitHere
    Case 3265 'item not found in this collection
        MsgBox "Query not found", vbOKOnly, "Temp Query Misssing"
         Resume ExitHere
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure LoadSQLfromTemporaryQuery of Module basSQLStatementsManager" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only

End Function