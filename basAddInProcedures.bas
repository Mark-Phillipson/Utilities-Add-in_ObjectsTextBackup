Option Compare Database

Option Explicit
Dim StringSetPropertiesObjectName As String

'
Public Sub TestMsgBox()
Dim c As ADODB.connection
Dim r As New ADODB.Recordset
Dim lngRecCounter As Long ' Used for Meter


On Error GoTo HandleErr




'Set the provider name
Set c = CurrentProject.connection
'Open a recordset with a keyset cursor
r.Open "Select * From [tmpTableNames]", c, adOpenKeyset, adLockPessimistic

If r.EOF Or r.BOF Then
Else
    r.MoveLast
    SysCmd acSysCmdInitMeter, "Walking through tmpTableNames", r.RecordCount
    r.MoveFirst
End If

Do Until r.EOF
    'r![id] = 0
    'r![TableName] = ""
    'r![Last Updated] = #9/9/2001#
    'r.Update
    lngRecCounter = lngRecCounter + 1
    SysCmd acSysCmdUpdateMeter, lngRecCounter
    r.MoveNext
Loop
r.Close
SysCmd acSysCmdRemoveMeter

ExitHere:
  Exit Sub

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 10 May 2004 09:13:29
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.TestMsgBox"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Sub

Public Function SelectAnObject()

On Error GoTo HandleErr

        DoCmd.OpenForm "frmOpenAnObject", acNormal

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 10 May 2004 09:13:29
HandleErr:
  Select Case Err.Number
  Case 2501 'The OpenForm action was canceled.
       Resume ExitHere
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.SelectAnObject"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function
Public Function OpenAdd_In_MainForm()

On Error GoTo HandleErr

    DoCmd.OpenForm "Navigation Form Utilities", acNormal

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 10 May 2004 09:13:29
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.OpenAdd_In_MainForm"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Function ChangeHistory(strProject As String)

On Error GoTo HandleErr

    DoCmd.OpenForm "frmChangeRequests", acNormal, , strProject

ExitHere:
  Exit Function

HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.ChangeHistory"
       Resume ExitHere
  End Select
  Resume 'Debug only
End Function


Public Function Test()
Dim db As DAO.database
Dim td As DAO.tableDef
Dim fld As DAO.field
Dim prp As DAO.Property
Dim prop1 As DAO.Property
Dim k As Integer
On Error GoTo HandleErr
Set db = codeDB
Set td = db.TableDefs("tmpMisspelledProperties")

For Each fld In td.Fields
    For Each prp In fld.Properties
        'Debug.Print "Name: " & prp.Name
        'Debug.Print "Type: " & prp.Type
        'Debug.Print "Inherited: " & prp.Inherited
        'Debug.Print "Value: " & prp.Value
        For Each prop1 In prp.Properties
            'Debug.Print "Name: " & prop1.Name
            'Debug.Print "Type: " & prop1.Type
            'Debug.Print "Inherited: " & prop1.Inherited
            'Debug.Print "Value: " & prop1.Value
        Next
    Next
Next
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2467, 3219, 3267, 3001, 3251 '
        Resume Next
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.test"  'ErrorHandler:$$N=basTestOnly.test
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function

Public Sub ListRelations(strTable As String, strDB As String, strNewDB As String)
    Dim lngLastCount As Long 'Added 10/05/2004
    Dim dbCode As DAO.database
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim rd As DAO.Recordset
    Dim lngRecCounter As Long
    Dim tdfLinked As tableDef

    On Error GoTo HandleErr

    Set dbCode = codeDB
    Set db = DBEngine.Workspaces(0).OpenDatabase(strDB)
    dbCode.Execute "Delete * From tmpTableNames;", dbFailOnError
    Set rs = dbCode.OpenRecordset("tmpTableNames")
    Set rd = dbCode.OpenRecordset("tmpTableNames")
    rs.AddNew
    rs![TableName] = strTable
    rs.Update
    AddRelatedTables db, strTable, rd
    rs.MoveLast
    lngRecCounter = rs.RecordCount
    Do
        rs.MoveFirst
        If lngRecCounter = lngLastCount Then Exit Do
        lngLastCount = lngRecCounter
        Do Until rs.EOF
            AddRelatedTables db, rs![TableName], rd
            rs.MoveNext
        Loop
        rs.MoveLast
        lngRecCounter = rs.RecordCount
    Loop
    db.Close
    On Error Resume Next
    Kill strNewDB
    On Error GoTo HandleErr
    Set db = DBEngine.Workspaces(0).CreateDatabase(strNewDB, dbLangGeneral)
    rs.MoveFirst
    lngRecCounter = 0
    SysCmd acSysCmdInitMeter, "Linking Tables...", rs.RecordCount
    Do Until rs.EOF
        Set tdfLinked = db.CreateTableDef(rs![TableName])
        tdfLinked.Attributes = dbAttachSavePWD
        tdfLinked.SourceTableName = rs![TableName]
        tdfLinked.Connect = ";DATABASE=" & strDB
        db.TableDefs.Append tdfLinked
        rs.MoveNext
        lngRecCounter = lngRecCounter + 1
        SysCmd acSysCmdUpdateMeter, lngRecCounter
    Loop
    SysCmd acSysCmdRemoveMeter
ExitHere:
  Exit Sub

HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.ListRelations"
       Resume ExitHere
  End Select
  Resume 'Debug only
End Sub

Public Function AddRelatedTables(db As DAO.database, ByVal strTable As String, rs As DAO.Recordset) As Boolean
    Dim rel As DAO.relation

On Error GoTo HandleErr

    AddRelatedTables = False
    rs.index = "TableName"
    For Each rel In db.Relations
        If rel.Table = strTable Then
            rs.MoveFirst
            rs.Seek "=", rel.ForeignTable
            If rs.NoMatch Then
                rs.AddNew
                rs![TableName] = rel.ForeignTable
                rs.Update
            End If
            AddRelatedTables = True
        End If
        If rel.ForeignTable = strTable Then
            rs.MoveFirst
            rs.Seek "=", rel.Table
            If rs.NoMatch Then
                rs.AddNew
                rs![TableName] = rel.Table
                rs.Update
            End If
            AddRelatedTables = True
        End If
    Next


ExitHere:
  Exit Function

HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basTestOnly.AddRelatedTables"
       Resume ExitHere
  End Select
  Resume 'Debug only
End Function


Public Function FillLinkedTables(strDBPath As String) As Boolean
    Dim strSQLtext  As String 'May 10
    Dim strTemp  As String '01/09/2009
    Dim intPosn As Integer '01/09/2009
    ' Fill tblLinkedTables with the linked table names of the strDBPath
    Dim p As String
    Dim fname As String
    Dim tmp As String
    Dim sz As Long
    Dim f_file As String
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim tdf As tableDef
    Dim fSystemObj As Boolean
    Dim fIsHidden  As Boolean
    Dim fShowSystem As Boolean
    Dim fShowHidden As Boolean
    Dim dbData As DAO.database
    Dim adhcTrackingHidden As Boolean
    Dim blnShow As Boolean
    Dim v As Variant
    Dim strDrive As String
    Dim strPath As String
    Dim strFileName As String
    Dim strExt As String

    On Error GoTo dbmGetTblNames_Err
    FillLinkedTables = True
    
    fShowSystem = False
    fShowHidden = False
    adhcTrackingHidden = False
    
    Set db = codeDB
    'tmp = "C:\My Documents\Work\ucp-address-data.mdb"
    'tmp = "\\AYNT4SVR1\Carrier\Data\ucp\ucp-address-data.mdb"
    Set dbData = DBEngine.Workspaces(0).OpenDatabase(strDBPath)
    
    Set rs = db.OpenRecordset("tblLinkedTables")
    
    'Set rs = db.OpenRecordset("tblLinkedTables", dbOpenTable)
    
    v = ZapTable(rs)
    
            dbData.TableDefs.Refresh
            For Each tdf In dbData.TableDefs
                'Debug.Print tdf.Name
                ' Check and see if this is a system object.
                fSystemObj = dbmisSystemObject(acTable, tdf.Name, tdf.Attributes)
                'If adhcTrackingHidden Then
                '    fIsHidden = IsHidden(tdf)
                'Else
                    ' If not tracking hidden objects,
                    ' just assume it's not hidden!
                    fIsHidden = False
                'End If
                ' Unless this is a system object and you're not showing system
                ' objects, or this table has its hidden bit set,
                ' add it to the list.
                If (fSystemObj Imp fShowSystem) And _
                 ((tdf.Attributes And dbHiddenObject) = 0) _
                 And (fIsHidden Imp fShowHidden) Then
                    ' only include linked tables
                    If tdf.Attributes = dbAttachedTable Then
                        If tdf.Name <> "Inbox" Then 'Dont include Outlook tables
                            rs.AddNew
                            rs![Table Name] = tdf.Name
                            rs![Server_Table_Name] = tdf.SourceTableName
                            intPosn = InStr(tdf.Connect, "DATABASE=") + 9
                            strPath = ParsePath(Mid(tdf.Connect, intPosn), 2)
                            strFileName = ParsePath(Mid(tdf.Connect, 11), 1)
                            If InStr(tdf.Connect, ".accdb") > 0 Then
                                strExt = ".accdb"
                            Else
                                strExt = ParsePath(Mid(tdf.Connect, 11), 4)
                            End If
                            rs![ConnectionString] = tdf.Connect
                            rs![Path] = strPath
                            rs![Type] = "Jet"
                            rs![Database Name] = strFileName & strExt
                            rs![ServerName] = "N/A"
                            rs.Update
                        End If
                    ElseIf tdf.Attributes = dbAttachedODBC Or tdf.Attributes = dbAttachedODBC + dbAttachSavePWD Then
                        rs.AddNew
                        rs![Table Name] = tdf.Name
                        rs![Server_Table_Name] = tdf.SourceTableName
                        rs![ConnectionString] = tdf.Connect
                        If InStr(tdf.Connect, "Oracle") > 0 Then
                            rs![Type] = "Oracle"
                            rs![Database Name] = ExtractConnectionValue("SERVER", tdf.Connect)
                            rs![ServerName] = ExtractConnectionValue("SERVER", tdf.Connect)
                        ElseIf InStr(tdf.Connect, "SQL Server") > 0 Or InStr(tdf.Connect, "SQL Native Client") > 0 Then
                            rs![Type] = "SQL Server"
                            rs![Database Name] = ExtractConnectionValue("DATABASE", tdf.Connect)
                            rs![ServerName] = ExtractConnectionValue("SERVER", tdf.Connect)
                        End If
                        rs.Update
                    End If
                End If
            Next tdf
    If Not rs.EOF Then rs.MoveLast
    rs.Close
    ' update subproject column
    strSQLtext = "UPDATE tblLinkedTables " & vbCrLf
    strSQLtext = strSQLtext & "  INNER JOIN tblTable_Sub_Project " & vbCrLf
    strSQLtext = strSQLtext & "          ON tblLinkedTables.[Table Name] = tblTable_Sub_Project.Table_Name SET tblLinkedTables.Sub_Project = [tblTable_Sub_Project]![Sub_Project];"
    db.Execute strSQLtext, dbFailOnError
    ' any subprojects that are empty set to not applicable
    strSQLtext = "UPDATE tblLinkedTables SET tblLinkedTables.Sub_Project = 'not applicable'" & vbCrLf
    strSQLtext = strSQLtext & "       WHERE (((tblLinkedTables.Sub_Project) Is Null));"
    db.Execute strSQLtext, dbFailOnError
    dbData.Close
    
    
dbmGetTblNames_Exit:
    Exit Function
dbmGetTblNames_Err:
    
    If Err.Number = 3044 Then
        MsgBox "Please enter a valid database path and name", vbExclamation
        FillLinkedTables = False
    Else
        MsgBox Err.Number & vbCrLf & vbCrLf & Err.Description
        FillLinkedTables = False
        GoTo dbmGetTblNames_Exit
    End If
    
    Resume
End Function


Public Sub FilltmpControls(strObjectName As String, ObjectType As Access.AcObjectType)
    Dim c As Control
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim f As Access.Form
    Dim r As Access.Report
    On Error GoTo HandleErr
    DoCmd.Close acForm, "frmControlsOnForm", acSaveYes
    Set db = codeDB
    Set rs = db.OpenRecordset("tmpControls")
    db.Execute "DELETE tmpControls.* FROM tmpControls;", dbFailOnError
    If ObjectType = acForm Then
        DoCmd.OpenForm strObjectName, acDesign
        Set f = Forms(strObjectName)
        For Each c In f.Controls
            rs.AddNew
            rs![Name] = c.Name
            rs![Parent] = c.Parent.Name
            rs![ControlType] = c.ControlType
            rs![Section] = c.Section
            On Error Resume Next
            rs![Italic] = c.FontItalic
            rs![Left] = c.Left
            On Error GoTo HandleErr
            rs.Update
        Next
        rs.Close
        
        DoCmd.OpenForm strObjectName, acDesign
        DoCmd.OpenForm "formselectcontrols", acNormal, , , acFormEdit, acDialog, strObjectName & "|" & ObjectType
    ElseIf ObjectType = acReport Then
        DoCmd.OpenReport strObjectName, acViewDesign
        Set r = Reports(strObjectName)
        For Each c In r.Controls
            rs.AddNew
            rs![Name] = c.Name
            'rs![Parent] = c.Parent.Name
            On Error Resume Next
            If c.ControlType <> acObjectFrame Then
                rs![Italic] = c.FontItalic
            End If
            On Error GoTo HandleErr
            rs![Left] = c.Left
            rs![ControlType] = c.ControlType
            rs![Section] = c.Section
            rs.Update
        Next
        rs.Close
        'DoCmd.OpenForm "frmControlsOnForm", acNormal, , , acFormEdit, acWindowNormal, strObjectName & "|" & ObjectType
        
        DoCmd.OpenForm "formselectcontrols", acNormal, , , acFormEdit, acDialog, strObjectName & "|" & ObjectType
    Else
        Exit Sub
    End If
    


ExitHere:
  Exit Sub

HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basAddInCode.FilltmpControls"
       Resume ExitHere
  End Select
  Resume 'Debug only
End Sub

Public Sub RenameControls(strFormName As String)
' Code Header inserted by VBA Code Commenter and Error Handler Add-In
'=============================================================
' basAddInProcedures.RenameControls
'-------------------------------------------------------------
' Purpose Rename all labels to lbl + Caption and all TextBoxes to txt + ControlSource
' Author : Mark Phillipson, 03 August 2006
' Notes :
'-------------------------------------------------------------
' Parameters
'-----------
' strFormName (String)
'-------------------------------------------------------------
' Revision History
'-------------------------------------------------------------
' 03 August 2006 MSP:
'=============================================================
' End Code Header block
    Dim strTemp As String 'Added 03/08/2006
    Dim f As Form
    Dim c As Control
    DoCmd.OpenForm strFormName, acDesign
    Set f = Forms(strFormName)
    For Each c In f.Controls
        If c.ControlType = acLabel Then
            strTemp = "lbl" & c.Caption
            If InStr(strTemp, ":") > 0 Then strTemp = Left(strTemp, InStr(strTemp, ":") - 1) & Mid(strTemp, InStr(strTemp, ":") + 1)
            If InStr(strTemp, "&") > 0 Then strTemp = Left(strTemp, InStr(strTemp, "&") - 1) & Mid(strTemp, InStr(strTemp, "&") + 1)
            c.Caption = Mid(strTemp, 4, 2)
        ElseIf c.ControlType = acTextBox Then
            strTemp = "txt" & c.ControlSource
        End If
        c.Name = strTemp
    Next
End Sub


Function EnterUK()
    Dim strDirection As String 'Added 04/10/2006
    Dim Title As String
    Dim msg As String, Defvalue As String
    Dim MyDB As database, MyRecords As DAO.Recordset, Total As Long
    Dim Entry$, answer As Variant, i As Long
    On Error GoTo HandleErr


    Dim CurrentType As Integer
    Dim CurrentName As String
    Set MyDB = CurrentDb
    CurrentType = Application.CurrentObjectType
    If CurrentType = acTable Or CurrentType = acQuery Then
        CurrentName = Application.CurrentObjectName
        On Error Resume Next
        Set MyRecords = MyDB.OpenRecordset(CurrentName)
        MyRecords.MoveLast
        Total = MyRecords.RecordCount
        If Total = -1 Then
            'Debug.Print "Count not available."
            Exit Function
        End If
        MyRecords.Close
        On Error GoTo HandleErr

    ElseIf CurrentType = A_FORM Then
        'DUMMY
    Else
        Exit Function
    End If




If Total > 1000 Or Total = 0 Then Total = 1000


DoCmd.OpenForm "frmEnterValueInColumn", acNormal, , , acFormEdit, acDialog

If Not IsLoaded("frmEnterValueInColumn") Then Exit Function
msg = "What do you wish to enter"  ' Set prompt.
Title = "Wack an expression in Field " & "(Max: " & str$(Total) & ")" ' Set title.
Defvalue = "UK"      ' Set default return value.
Entry$ = Nz(Forms![frmEnterValueInColumn].[txtValue], "")
If Entry$ = "" Then Entry$ = " "


msg = "Enter Number of Fields to Enter " & vbCrLf & Entry$ & vbCrLf & "in"  ' Set prompt.
Title = "Enter " & Entry$ & " in Column " & "(Max: " & str$(Total) & ")" ' Set title.
Defvalue = "10"      ' Set default return value.
answer = Forms![frmEnterValueInColumn].[txtQty]

'"{DOWN}"
strDirection = "{" & UCase(Forms![frmEnterValueInColumn].[cboDirection]) & "}"
If strDirection = "{LEFT}" Then strDirection = "+{Tab}"
DoCmd.Close acForm, "frmEnterValueInColumn", acSaveYes


If Val(answer) < 1 Then Exit Function

If Val(answer) > 1000 Then
    MsgBox "Sorry max reached, an update query would be more effecient", vbExclamation, "More than 1,000"
    Exit Function
End If
    'Put braces round any brackets as sendkeys won't send them as is
    Entry$ = Replace(Entry$, "(", "{(}")
    Entry$ = Replace(Entry$, ")", "{)}")
    
   For i = 1 To Val(answer)    ' Set up counting loop.
       SendKeys Entry$ & strDirection, True
   Next i

    
    
    SendKeys "{NUMLOCK}", True



ExitHere:
    Exit Function

' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error
HandleErr:
    Select Case Err.Number
    'Case # '
    ' MsgBox "", vbExclamation, ""
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected Error in basImportFiles.EnterUK" 'ErrorHandler:$$N=basImportFiles.EnterUK
        Resume ExitHere
    End Select
Resume
End Function


Public Function ExtractConnectionValue(strValueName As String, strConnStr As String)
    Dim strTemp  As String '01/09/2009
    Dim strResult As String '01/09/2009
    Dim intPosn As Integer '01/09/2009
    intPosn = InStr(strConnStr, strValueName & "=") + (Len(strValueName) + 1)
    strTemp = Mid(strConnStr, intPosn)
    If InStr(strTemp, ";") > 0 Then
        strResult = Left(strTemp, InStr(strTemp, ";") - 1)
    Else
        strResult = strTemp
    End If
    ExtractConnectionValue = strResult
End Function

Public Function SelectControlsOnForm()
    Dim f As Access.Form
    Dim f1 As Access.Form
 If Application.CurrentObjectType = acForm Then
     FilltmpControls Application.CurrentObjectName, acForm
    
 ElseIf Application.CurrentObjectType = acReport Then
    FilltmpControls Application.CurrentObjectName, acReport
 ElseIf Application.CurrentObjectType = acTable Then
    Set f = Screen.ActiveDatasheet
    DoCmd.OpenForm "frmGetColumnName", acNormal, , , acFormEdit, acDialog

    If IsLoaded("frmGetColumnName") Then
        Set f1 = Forms![frmGetColumnName]
        If f1.chkLinked Then
            f.SelLeft = f1.txtColumnNumer
        Else
           f.SelLeft = f1.txtColumnNumer + 1
            
        End If
        DoCmd.Close acForm, "frmGetColumnName"
    End If
 End If
 
End Function

Public Function ShowDebugging()
    'Test function to show debugging in action using voice recognition
    Dim k As Integer
    For k = 1 To 5
        Beep
        Debug.Print "number of times through the loop; " & k
    Next
    End
End Function
Public Function InsertSpaces(stringValue As String) As String
    Dim stringResult As String 'Jul 12
    Dim booleanLastUpperCase As Boolean
    Dim stringNextLetter As String
    On Error GoTo InsertSpaces_Error

    If Len(stringValue) = 0 Then
        Exit Function
    End If
    Dim stringLetter As String
    Dim k As Integer
    For k = 1 To Len(stringValue)
        stringLetter = Mid(stringValue, k, 1)
        If Not k = Len(stringValue) Then
            stringNextLetter = Mid(stringValue, k + 1, 1)
        Else
            stringNextLetter = ""
        End If
        
        If Asc(stringLetter) = Asc(UCase(stringLetter)) Then
            If Len(stringNextLetter) > 0 And stringLetter <> "&" Then
                If Not booleanLastUpperCase Or Not Asc(stringNextLetter) = Asc(UCase(stringNextLetter)) Then
                    stringResult = stringResult & " " & stringLetter
                Else
                    stringResult = stringResult & stringLetter
                End If
            Else
                stringResult = stringResult & stringLetter
            End If
            
            booleanLastUpperCase = True
        Else
            stringResult = stringResult & stringLetter
            booleanLastUpperCase = False
        End If


    Next
    InsertSpaces = Trim(stringResult)

ExitHere:
   Exit Function

InsertSpaces_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure InsertSpaces of Module basAddInProcedures" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only

End Function


Public Function ListForms()
    Dim stringIgnore As String
      With COMAddIns("VBACodeWriter.Connect")
         'Make sure the COM add-in is loaded.
        .Connect = True
        
        stringIgnore = .object.ListForms(False)
    End With

                                         
End Function

Public Function ListControls()
 Dim stringIgnore As String
      With COMAddIns("VBACodeWriter.Connect")
         'Make sure the COM add-in is loaded.
        .Connect = True
        
        stringIgnore = .object.ListControls(False)
    End With

End Function

Public Function SetBooleanProperty(stringPropertyName As String)
Dim currentObject As Object
Dim activeControl As Access.Control

    On Error Resume Next
    Dim stringObject As String
    stringObject = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Err.Clear
        stringObject = Screen.ActiveReport.Name
    Else
        Set currentObject = Forms(stringObject)
    End If
    If Not Err.Number = 0 Then
        Exit Function
    Else
        Set currentObject = Reports(stringObject)
    End If
    On Error GoTo SetBooleanProperty_Error
    If Len(stringObject) > 0 Then
        
        For Each activeControl In currentObject.Controls
            If activeControl.InSelection Then
                Select Case stringPropertyName
                Case "Visible"
                    activeControl.Visible = Not activeControl.Visible
                Case "Enabled"
                    activeControl.Enabled = Not activeControl.Enabled
                Case "Vertical"
                    activeControl.Vertical = Not activeControl.Vertical
                Case "Locked"
                    activeControl.Locked = Not activeControl.Locked
                Case "TabStop"
                    activeControl.TabStop = Not activeControl.TabStop
                End Select
            End If
        Next
    End If

ExitHere:
   Exit Function

SetBooleanProperty_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SetBooleanProperty of Module basAddInProcedures" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function

Public Function RenameActiveControl()
Dim currentObject As Object
Dim activeControl As Access.Control
    Dim currentDatabase As DAO.database
    Dim Recordset As DAO.Recordset
    Dim stringSQLText As String
    On Error GoTo RenameActiveControl_Error

    Set currentDatabase = codeDB
    
    On Error Resume Next
    Dim stringObject As String
    stringObject = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Err.Clear
        stringObject = Screen.ActiveReport.Name
        If Not Err.Number = 0 Then
            Exit Function
        Else
            Set currentObject = Reports(stringObject)
        End If
    Else
        Set currentObject = Forms(stringObject)
    End If
    On Error GoTo RenameActiveControl_Error
    Dim stringNewName As String
    Dim stringSuffix As String
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            stringSQLText = "SELECT tblRVBAAccessObjTags.[Type No]" & vbCrLf
            stringSQLText = stringSQLText & "           , tblRVBAAccessObjTags.Tag" & vbCrLf
            stringSQLText = stringSQLText & "        FROM tblRVBAAccessObjTags" & vbCrLf
            stringSQLText = stringSQLText & "       WHERE (((tblRVBAAccessObjTags.[Type No])=" & activeControl.ControlType & "));"
            Set Recordset = currentDatabase.OpenRecordset(stringSQLText, dbOpenSnapshot)
            If Not Recordset.EOF Then
                Recordset.MoveFirst
                stringSuffix = Recordset![tag]
            End If
            Dim stringOldName As String
            stringOldName = activeControl.Name
            If Len(stringSuffix & "") > 0 Then
                RemoveOldTags activeControl.ControlType, stringOldName, "TextBox", acTextBox
                RemoveOldTags activeControl.ControlType, stringOldName, "txt", acTextBox
                RemoveOldTags activeControl.ControlType, stringOldName, "ComboBox", acComboBox
                RemoveOldTags activeControl.ControlType, stringOldName, "cbo", acComboBox
                RemoveOldTags activeControl.ControlType, stringOldName, "Label", acLabel
                RemoveOldTags activeControl.ControlType, stringOldName, "lbl", acLabel
                
                If InStr(stringOldName, stringSuffix) = 0 Then
                    stringNewName = stringOldName & stringSuffix
                    activeControl.Name = stringNewName
                End If
            End If
        End If
    Next

ExitHere:
   Exit Function

RenameActiveControl_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure RenameActiveControl of Module basAddInProcedures" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only

End Function

Public Function SetGridlinesStyles()
    Dim currentObject As Object
    Dim activeControl As Access.Control
    Dim stringObject As String
    On Error Resume Next
    stringObject = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Err.Clear
        stringObject = Screen.ActiveReport.Name
        If Not Err.Number = 0 Then
            Exit Function
        Else
            Set currentObject = Reports(stringObject)
        End If
    Else
        Set currentObject = Forms(stringObject)
    End If
    On Error GoTo SetGridlinesStyles_Error
    DoCmd.OpenForm "Set Gridlines Style", acNormal, , , acFormEdit, acDialog
    If Not IsLoaded("Set Gridlines Style") Then Exit Function
    Dim ExampleForm As Access.Form
    Set ExampleForm = Forms![Set Gridlines Style]
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            With ExampleForm
                activeControl.GridlineStyleBottom = .ExampleTextBox.GridlineStyleBottom
                activeControl.GridlineStyleTop = .ExampleTextBox.GridlineStyleTop
                activeControl.GridlineStyleLeft = .ExampleTextBox.GridlineStyleLeft
                activeControl.GridlineStyleRight = .ExampleTextBox.GridlineStyleRight
            End With
        End If
    Next
    DoCmd.Close acForm, ExampleForm.Name, acSaveYes

ExitHere:
   Exit Function

SetGridlinesStyles_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SetGridlinesStyles of Module basAddInProcedures" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function

Public Sub TestMenu()


End Sub

Public Function SizeandPosition(stringMode As String)
    Dim currentObject As Object
    Dim activeControl As Access.Control
    Dim stringObject As String
    


    On Error GoTo SizeandPosition_Error

    stringObject = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Err.Clear
        stringObject = Screen.ActiveReport.Name
        If Not Err.Number = 0 Then
            Exit Function
        Else
            Set currentObject = Reports(stringObject)
        End If
    Else
        Set currentObject = Forms(stringObject)
    End If
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            If stringMode = "Copy" Then
                Application.TempVars.Add "TopPosition", activeControl.Top
                Application.TempVars.Add "LeftPosition", activeControl.Left
                Application.TempVars.Add "Height", activeControl.Height
                Application.TempVars.Add "Width", activeControl.Width
                Exit For
            ElseIf stringMode = "Paste" Then
                activeControl.Top = Application.TempVars![TopPosition]
                activeControl.Left = Application.TempVars![LeftPosition]
                activeControl.Height = Application.TempVars![Height]
                activeControl.Width = Application.TempVars![Width]
            End If
        End If
    Next


ExitHere:
   Exit Function

SizeandPosition_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SizeandPosition of Module basAddInProcedures" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function


Public Function ChangePropertiesCustom()
'    If Len(StringSetPropertiesObjectName) > 0 And StringSetPropertiesObjectName <> Application.CurrentObjectName Then
        ' this is a different object so close the form and build from scratch
        DoCmd.Close acForm, "Select Properties", acSaveYes
'    End If
    StringSetPropertiesObjectName = Application.CurrentObjectName
    If Application.CurrentObjectType = acForm Then
    
        DoCmd.OpenForm Application.CurrentObjectName, acDesign
        DoCmd.OpenForm "Select Properties", acNormal, , , acFormEdit, acWindowNormal, Application.CurrentObjectName & "|" & Application.CurrentObjectType
    ElseIf Application.CurrentObjectType = acReport Then
        DoCmd.OpenReport Application.CurrentObjectName, acViewDesign
        DoCmd.OpenForm "Select Properties", acNormal, , , acFormEdit, acWindowNormal, Application.CurrentObjectName & "|" & Application.CurrentObjectType
    End If
                

End Function

Public Function SetComboBoxProperties()

    Dim currentObject As Object
    Dim activeControl As Access.Control
    Dim stringObject As String
    Dim PropertiesForm As Form
    On Error Resume Next
    stringObject = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Err.Clear
        stringObject = Screen.ActiveReport.Name
        If Not Err.Number = 0 Then
            Exit Function
        Else
            Set currentObject = Reports(stringObject)
        End If
    Else
        Set currentObject = Forms(stringObject)
    End If
    On Error GoTo SetComboBoxProperties_Error
    On Error Resume Next
    DoCmd.OpenForm "Set ComboBox Properties", acNormal, , , , acHidden
    Dim stringPropertiesFormName As String
    If Not Err.Number = 0 Then
        DoCmd.OpenForm " Set ComboBox Properties Template", acNormal, , , , acHidden
        Set PropertiesForm = Forms![Set ComboBox Properties Template]
        stringPropertiesFormName = "Set ComboBox Properties Template"
    Else
        Set PropertiesForm = Forms![Set ComboBox Properties]
        stringPropertiesFormName = "Set ComboBox Properties"
    End If
    On Error GoTo SetComboBoxProperties_Error
    ' initialise the properties from the selectedComboBox
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            With activeControl
                Select Case .ControlType
                Case acComboBox, acListBox
                    PropertiesForm.ColumnHeadsCheckBox = .ColumnHeads
                    PropertiesForm.RowSourceTypeComboBox = .RowSourceType
                    PropertiesForm.SQLTextTextBox = .RowSource
                    PropertiesForm.ColumnCountTextBox = .ColumnCount
                    If Len(.ColumnWidths & "") > 0 Then
                        PropertiesForm.ColumnWidthsTextBox = ConvertColumnWidths(.ColumnWidths, ToCentimetres, 0)
                    End If
                End Select
            End With
        End If
    Next
    PropertiesForm.Visible = True
    PropertiesForm.SetFocus
    Do
        DoEvents
        If IsLoaded(stringPropertiesFormName) Then
            If PropertiesForm.Visible = False Then
                Exit Do
            End If
        Else
            Exit Function
        End If
    Loop
    Dim integerTotalWidth As Integer
    If Not IsLoaded(stringPropertiesFormName) Then Exit Function
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            With PropertiesForm
                Select Case activeControl.ControlType
                Case acComboBox, acListBox
                    activeControl.ColumnHeads = PropertiesForm.ColumnHeadsCheckBox
                    activeControl.RowSourceType = PropertiesForm.RowSourceTypeComboBox
                    activeControl.RowSource = PropertiesForm.SQLTextTextBox
                    activeControl.ColumnCount = Nz(PropertiesForm.ColumnCountTextBox, 1)
                    If Len(PropertiesForm.ColumnWidthsTextBox & "") > 0 Then
                        activeControl.ColumnWidths = Nz(ConvertColumnWidths(PropertiesForm.ColumnWidthsTextBox, totwips, integerTotalWidth), "")
                        If Not activeControl.ControlType = acListBox Then
                            activeControl.ListWidth = integerTotalWidth
                        End If
                        
                    End If
                End Select
            End With
        End If
    Next
    DoCmd.Close acForm, PropertiesForm.Name, acSaveYes

ExitHere:
    Exit Function

SetComboBoxProperties_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SetComboBoxProperties of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Public Function SetLookupProperties(blnColumnHeads As Boolean, stringRowSourceType As String, stringRowSource As String, integerColumnCount As Integer, _
stringColumnWidths As String, integerListWidth As Integer)

    Dim PropertiesForm As Form
    On Error Resume Next
    DoCmd.OpenForm "Set ComboBox Properties", acNormal, , , , acHidden
    Dim stringFormName As String
    If Not Err.Number = 0 Then
        DoCmd.OpenForm " Set ComboBox Properties Template", acNormal, , , , acHidden
        Set PropertiesForm = Forms![Set ComboBox Properties Template]
        stringFormName = "Set ComboBox Properties Template"
    Else
        Set PropertiesForm = Forms![Set ComboBox Properties]
        stringFormName = "Set ComboBox Properties"
    End If
    PropertiesForm.ColumnHeadsCheckBox = blnColumnHeads
    PropertiesForm.RowSourceTypeComboBox = stringRowSourceType
    PropertiesForm.SQLTextTextBox = stringRowSource
    PropertiesForm.ColumnCountTextBox = integerColumnCount
    If Len(stringColumnWidths & "") > 0 Then
        PropertiesForm.ColumnWidthsTextBox = ConvertColumnWidths(stringColumnWidths, ToCentimetres, 0)
    End If
    PropertiesForm.Visible = True
    PropertiesForm.SetFocus
    Do
        DoEvents
        If IsLoaded(stringFormName) Then
            If PropertiesForm.Visible = False Then
                Exit Do
            End If
        Else
            Exit Function
        End If
    Loop
    Dim integerTotalWidth As Integer
    If Not IsLoaded(stringFormName) Then Exit Function
    With PropertiesForm
            blnColumnHeads = PropertiesForm.ColumnHeadsCheckBox
            stringRowSourceType = PropertiesForm.RowSourceTypeComboBox
            stringRowSource = PropertiesForm.SQLTextTextBox
            integerColumnCount = Nz(PropertiesForm.ColumnCountTextBox, 1)
            If Len(PropertiesForm.ColumnWidthsTextBox & "") > 0 Then
                stringColumnWidths = Nz(ConvertColumnWidths(PropertiesForm.ColumnWidthsTextBox, totwips, integerTotalWidth), "")
                integerListWidth = integerTotalWidth
            End If
    End With
    DoCmd.Close acForm, stringFormName, acSaveYes

ExitHere:
   Exit Function

SetLookupProperties_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SetLookupProperties of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Private Function ConvertColumnWidths(stringColumnWidths As String, conversion As MeasurementConversion, integerTotalWidth As Integer) As String
    Dim variantArray As Variant
    On Error GoTo ConvertColumnWidths_Error
    variantArray = ConvertDelimitedToArray(";", stringColumnWidths)
    Dim integerCounter As Integer
    Dim singleWidth As Single
    Dim stringTemporary As String
    For integerCounter = 0 To UBound(variantArray) - 1
        If Len(variantArray(integerCounter) & "") > 0 Then
            variantArray(integerCounter) = Replace(variantArray(integerCounter), "cm", "")
            singleWidth = Round(ConvertMeasurement(conversion, variantArray(integerCounter)), 2)
            If conversion = totwips Then
                integerTotalWidth = integerTotalWidth + singleWidth
            End If
            stringTemporary = stringTemporary & IIf(Len(stringTemporary) > 0, ";", "") & Trim(CStr(singleWidth)) & IIf(conversion = ToCentimetres, "cm", "")
        End If
    Next
    ConvertColumnWidths = stringTemporary


ExitHere:
   Exit Function

ConvertColumnWidths_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure ConvertColumnWidths of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Public Function ConvertMeasurement(conversion As MeasurementConversion, variantvalue As Variant) As Variant
    If Len(variantvalue & "") = 0 Then
        ConvertMeasurement = variantvalue
        Exit Function
    End If
    If conversion = ToCentimetres Then
        ConvertMeasurement = variantvalue / 567
    ElseIf conversion = totwips Then
        ConvertMeasurement = 567 * variantvalue
    End If

End Function

Public Function GetLatestExtrasMacro(Optional blnFromAddin As Boolean = False)
    Dim stringSourceDatabase As String
    On Error GoTo GetLatestExtrasMacro_Error
    stringSourceDatabase = "C:\MSOffice\access\Utilities Add-in.accda"
    If Not CurrentDb.Name = stringSourceDatabase Then
        'On Error Resume Next
        'DoCmd.DeleteObject acForm, "Set ComboBox Properties"
        On Error GoTo GetLatestExtrasMacro_Error
        DoCmd.TransferDatabase acImport, "Microsoft Access", stringSourceDatabase, acForm, "Set ComboBox Properties Template" _
        , "Set ComboBox Properties"
        On Error Resume Next
        DoCmd.Rename "Set ComboBox Properties", acForm, "Set ComboBox Properties1"
        'On Error Resume Next
        'DoCmd.DeleteObject acMacro, "ExtrasRibbon"
        On Error GoTo GetLatestExtrasMacro_Error
        DoCmd.TransferDatabase acImport, "Microsoft Access", stringSourceDatabase, acMacro, "ExtrasRibbon", "ExtrasRibbon"
        
        On Error Resume Next
        DoCmd.Rename "ExtrasRibbon", acMacro, "ExtrasRibbon1"
        On Error GoTo GetLatestExtrasMacro_Error
        DoCmd.TransferDatabase acImport, "Microsoft Access", stringSourceDatabase, acMacro, "AutoKeys", "AutoKeys"
        On Error Resume Next
        DoCmd.Rename "AutoKeys", acMacro, "AutoKeys1"
        On Error GoTo GetLatestExtrasMacro_Error
        If blnFromAddin Then
            ManageReference True
        End If
    End If

ExitHere:
   Exit Function

GetLatestExtrasMacro_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure GetLatestExtrasMacro of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function

Private Sub RemoveOldTags(ControlType As AcControlType, stringName As String, stringTag As String, NotControlType As AcControlType)
    If Not ControlType = NotControlType Then
        If Left(stringName, Len(stringTag)) = stringTag Or Right(stringName, Len(stringTag)) = stringTag Then
            stringName = Replace(stringName, stringTag, "")
        End If
    End If

End Sub

Public Function OpenManagedLinkedTables()
    On Error GoTo OpenManagedLinkedTables_Error
    DoCmd.OpenForm "frmManageLinkedTables", acNormal

ExitHere:
   Exit Function

OpenManagedLinkedTables_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure OpenManagedLinkedTables of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function

Public Sub ShowBrowseSchema(stringTableName As String)
    Dim stringSourceTableName As String
        Dim codeDatabase As DAO.database
        Dim RecordsetTemporary As DAO.Recordset
        Dim stringSQLText As String
        On Error GoTo ShowBrowseSchema_Error
        DoCmd.Close acForm, "Browse Schema", acSaveYes
        Set codeDatabase = codeDB
        stringSQLText = "SELECT [Temporary Table Schema].*" & vbCrLf
        stringSQLText = stringSQLText & "        FROM [Temporary Table Schema];"
        Set RecordsetTemporary = codeDatabase.OpenRecordset(stringSQLText)
        ZapTable RecordsetTemporary
        Dim currentDatabase As DAO.database
        Dim Recordset As DAO.Recordset
        Set currentDatabase = CurrentDb
        Dim tableDefinition As DAO.tableDef
        Set tableDefinition = currentDatabase.TableDefs(stringTableName)
        Dim field As DAO.field
        Dim blnLinked As Boolean
        Dim blnAccessDatabase As Boolean
        Dim externalDatabase As DAO.database
        Dim stringDatabase As String
        Dim blnSQLServer As Boolean
        If Len(tableDefinition.Connect & "") > 0 Then
            blnLinked = True
            If InStr(tableDefinition.Connect, "DATABASE=") > 0 And InStr(tableDefinition.Connect, "ODBC") = 0 Then
                Application.TempVars.Add "TableType", "Access Linked"
                blnAccessDatabase = True
                stringDatabase = Mid(tableDefinition.Connect, InStr(tableDefinition.Connect, "DATABASE=") + 9)
                stringSourceTableName = tableDefinition.SourceTableName
                Application.TempVars.Add "SourceTableName", tableDefinition.SourceTableName
                Application.TempVars.Add "ODBCConnectionString", tableDefinition.Connect
                Set externalDatabase = DBEngine(0).OpenDatabase(stringDatabase)
                Set tableDefinition = externalDatabase.TableDefs(stringSourceTableName)
            ElseIf InStr(tableDefinition.Connect, "SQL Server") > 0 And InStr(tableDefinition.Connect, "ODBC") > 0 Then
                Application.TempVars.Add "TableType", "SQL Server"
                blnSQLServer = True
                stringDatabase = ""
                Application.TempVars.Add "SQLServerConnectionString", GetADOConnectionString(tableDefinition.Connect)
                Application.TempVars.Add "ODBCConnectionString", tableDefinition.Connect
            Else
                Application.TempVars.Add "TableType", "Other Linked"
                blnAccessDatabase = False
                stringDatabase = ""
            End If
        Else
            Application.TempVars.Add "TableType", "Access"
            blnLinked = False
        End If
     For Each field In tableDefinition.Fields
        RecordsetTemporary.AddNew
        RecordsetTemporary![FieldName] = field.Name
        RecordsetTemporary![OldFieldName] = field.Name
        RecordsetTemporary![DataType] = FieldType(field.Type)
        RecordsetTemporary![FieldSize] = field.Size
        RecordsetTemporary![Required] = field.Required
        RecordsetTemporary![AllowZeroLength] = field.AllowZeroLength
        If Application.TempVars![TableType] = "SQL Server" And InStr(Application.TempVars![SQLServerConnectionString], "windows.net") = 0 Then
            RecordsetTemporary![DefaultValue] = GetDefaultValueSQLServer(Application.TempVars![SQLServerConnectionString], stringTableName, field.Name)
        Else
            RecordsetTemporary![DefaultValue] = field.DefaultValue
        End If
        If HasProperty(field, "Caption") Then
            RecordsetTemporary![Caption] = field.Properties("Caption")
        End If
        If HasProperty(field, "Description") Then
            RecordsetTemporary![Description] = field.Properties("Description")
        End If
        If HasProperty(field, "Format") Then
            RecordsetTemporary![Format] = field.Properties("Format")
        End If
        RecordsetTemporary.Update
    Next
    RecordsetTemporary.Close
    DoCmd.OpenForm "Browse Schema", acNormal
    With Forms![Browse Schema]
        .TableNameTextBox = stringTableName
        .DatabaseTextBox = stringDatabase
        If blnLinked And Not blnAccessDatabase And Not blnSQLServer Then
            .CaptionTextBox.SetFocus
            .FieldNameTextBox.Enabled = False
            .RequiredCheckBox.Enabled = False
            .DeleteButton.Enabled = False
            .AllowZeroLengthCheckBox.Enabled = False
            .CreateFieldButton.Enabled = False
            .DefaultValueTextBox.Enabled = False
        ElseIf blnAccessDatabase Then
            .FieldNameTextBox.Enabled = True
            .RequiredCheckBox.Enabled = True
            .DeleteButton.Enabled = True
            .AllowZeroLengthCheckBox.Enabled = True
            .CreateFieldButton.Enabled = True
            .DefaultValueTextBox.Enabled = True
        ElseIf blnSQLServer Then
            .FieldNameTextBox.Enabled = False
            .RequiredCheckBox.Enabled = True
            .DeleteButton.Enabled = True
            .AllowZeroLengthCheckBox.Enabled = False
            .CreateFieldButton.Enabled = True
            .DefaultValueTextBox.Enabled = False
        End If
    End With


ExitHere:
   Exit Sub

ShowBrowseSchema_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure ShowBrowseSchema of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Sub

Public Function GetADOConnectionString(stringODBCConnectionString As String) As String
    'ODBC;DRIVER=sql server;SERVER=USER-HP;Trusted_Connection=Yes;APP=Microsoft Office 2010;DATABASE=MyDatabase
    'ODBC;DRIVER=SQL Server Native Client 11.0;SERVER=f5ath01b5q.database.windows.net,1433;APP=Microsoft Office 2010;DATABASE=LightswA4vZIOuXz;
    Dim AccessForm As Access.Form
    Dim stringServer As String
    If InStr(stringODBCConnectionString, "SERVER=") > 0 Then
        stringServer = Mid(stringODBCConnectionString, InStr(stringODBCConnectionString, "SERVER=") + 7)
    End If
    If InStr(stringServer, ";") > 0 Then
        stringServer = Left(stringServer, InStr(stringServer, ";") - 1)
    End If
    Dim stringDatabase As String
    If InStr(stringODBCConnectionString, "DATABASE=") > 0 Then
        stringDatabase = Mid(stringODBCConnectionString, InStr(stringODBCConnectionString, "DATABASE=") + 9)
    End If
    If InStr(stringODBCConnectionString, "Trusted_Connection=Yes") > 0 Then
        GetADOConnectionString = "Provider=SQLOLEDB;Data Source=" & stringServer & ";Initial Catalog=" & stringDatabase & ";Trusted_Connection=Yes;"
    Else
        Dim stringFormName As String: stringFormName = "Login"
        If Not IsLoaded(stringFormName) Then
            DoCmd.Close acForm, stringFormName, acSaveYes
            DoCmd.OpenForm stringFormName, acNormal, , , acFormEdit, acDialog
        End If
        If Not IsLoaded(stringFormName) Then Exit Function
        Set AccessForm = Forms(stringFormName)
        GetADOConnectionString = "Provider=SQLOLEDB;Data Source=" & stringServer & ";Initial Catalog=" & stringDatabase _
        & ";UID=" & AccessForm![UsernameTextBox] & ";PWD=" & AccessForm![PasswordTextBox]
    End If
        
End Function
Private Function GetObjectName(AccessObject As AcObjectType, Optional stringArguments As String = "") As String
    DoCmd.Close acForm, "Generic Access Object Picker", acSaveYes
    DoCmd.OpenForm "Generic Access Object Picker", acNormal, , , acFormEdit, acDialog, IIf(Len(stringArguments) > 0, stringArguments, AccessObject)
    If IsLoaded("Generic Access Object Picker") Then
        GetObjectName = Forms![Generic Access Object Picker].ResultsListBox
    End If

End Function


Public Function GetFormName()
    On Error Resume Next
    Screen.activeControl = GetObjectName(acForm)
    On Error GoTo 0
End Function

Public Function SaveFormAsText()
    On Error GoTo SaveFormAsText_Error
    Dim stringFormName As String
    stringFormName = GetObjectName(acForm)
    Dim stringFolderName  As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Filters.Clear
        .InitialView = msoFileDialogViewList
        .Title = "Select the Folder where the form text file should be saved"
        If .Show Then
            stringFolderName = .SelectedItems(1)
        End If
    End With
    
    Application.SaveAsText acForm, stringFormName, stringFolderName & "\" & stringFormName & ".txt"

ExitHere:
   Exit Function

SaveFormAsText_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SaveFormAsText of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

    
End Function

Public Function LoadFormFromText()
    Dim stringFilename  As String
    On Error GoTo LoadFormFromText_Error
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "Text File", "*.txt"
        .Filters.Add "All Files", "*.*"
        .InitialView = msoFileDialogViewList
        .Title = "Select text file"
        If .Show Then
            stringFilename = .SelectedItems(1)
        End If
    End With
    If Len(stringFilename) = 0 Then Exit Function
    Dim stringFormName As String
    stringFormName = Mid(stringFilename, InStrRev(stringFilename, "\") + 1)
    stringFormName = Left(stringFormName, Len(stringFormName) - 4)
    Application.LoadFromText acForm, stringFormName, stringFilename

ExitHere:
   Exit Function

LoadFormFromText_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure LoadFormFromText of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function

Public Function GetReportName()
    On Error Resume Next
    Screen.activeControl = GetObjectName(acReport)
    On Error GoTo 0
End Function

Public Function GetTableName()
    On Error Resume Next
    Screen.activeControl = GetObjectName(acTable)
    On Error GoTo 0
End Function

Public Function GetTableFieldName()
    Dim stringTableName As String
    stringTableName = GetObjectName(acTable)
    If Len(stringTableName & "") = 0 Then
        Exit Function
    End If
    On Error Resume Next
    Screen.activeControl = GetObjectName(acTable, "Table|" & stringTableName)
    On Error GoTo 0
End Function

Public Function GetQueryFieldName()
    Dim stringQueryName As String
     stringQueryName = GetObjectName(acQuery)
    If Len(stringQueryName & "") = 0 Then
        Exit Function
    End If
    On Error Resume Next
    Screen.activeControl = GetObjectName(acQuery, "Query|" & stringQueryName)
    On Error GoTo 0
End Function

Public Function GetQueryName()
    On Error Resume Next
    Screen.activeControl = GetObjectName(acQuery)
    On Error GoTo 0
End Function

Public Function BrowseAndCreateSelectQuery(Optional AccessObject As AcObjectType = acTable)

     Dim stringTableName As String
     On Error GoTo BrowseAndCreateSelectQuery_Error
     stringTableName = GetObjectName(AccessObject)
     CreateSelectQuery stringTableName

ExitHere:
   Exit Function

BrowseAndCreateSelectQuery_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure BrowseAndCreateSelectQuery of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Public Function SetCommandButtonProperties()
    Dim currentObject As Object
    Dim blnButtonSelected As Boolean
    Dim activeControl As Access.Control
    Dim stringObject As String
    Dim dialogForm As Access.Form
    On Error Resume Next
    stringObject = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Exit Function
    End If
    Set currentObject = Forms(stringObject)
    On Error GoTo SetCommandButtonProperties_Error
    Dim stringDialogFormName As String: stringDialogFormName = "Create and Change Command Button Properties"
    DoCmd.OpenForm stringDialogFormName, acNormal, , , acFormEdit, acHidden
    Set dialogForm = Forms(stringDialogFormName)
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            With activeControl
                Select Case .ControlType
                Case acCommandButton
                    blnButtonSelected = True
                    dialogForm.CaptionTextBox = .Caption
                    dialogForm.CommandButtonNameTextBox = .Name
                    dialogForm.Transparentcheckbox = .Transparent
                End Select
            End With
        End If
    Next
    If Not blnButtonSelected Then
        For Each activeControl In currentObject.Controls
            activeControl.InSelection = False
        Next
        Set activeControl = Application.CreateControl(currentObject.Name, acCommandButton, acDetail, "")
        activeControl.InSelection = True
    End If
    dialogForm.Visible = True
    dialogForm.SetFocus
    Do
        DoEvents
        If IsLoaded(stringDialogFormName) Then
            If dialogForm.Visible = False Then
                Exit Do
            End If
        Else
            Exit Function
        End If
    Loop
    If Not IsLoaded(stringDialogFormName) Then Exit Function
    For Each activeControl In currentObject.Controls
        If activeControl.InSelection Then
            With dialogForm
                Select Case activeControl.ControlType
                Case acCommandButton
                    activeControl.Caption = .CaptionTextBox
                    activeControl.Name = .CommandButtonNameTextBox
                    activeControl.Transparent = .Transparentcheckbox
                End Select
            End With
        End If
    Next
    DoCmd.Close acForm, stringDialogFormName, acSaveYes
        

ExitHere:
   Exit Function

SetCommandButtonProperties_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SetCommandButtonProperties of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

    
End Function

Public Function SetFormProperties()
    Dim CurrentForm As Access.Form
    Dim stringFormName As String
    Dim dialogForm As Access.Form
    On Error Resume Next
    stringFormName = Screen.ActiveForm.Name
    If Not Err.Number = 0 Then
        Exit Function
    End If
    On Error GoTo SetFormProperties_Error
    Set CurrentForm = Forms(stringFormName)
    Dim stringDialogFormName As String: stringDialogFormName = "Set Form Properties"
    DoCmd.OpenForm stringDialogFormName, acNormal, , , acFormEdit, acHidden
    Set dialogForm = Forms(stringDialogFormName)
    With dialogForm
        .CaptionTextBox = CurrentForm.Caption
        .Caption = "Set Form Properties for " & CurrentForm.Name
        .DefaultViewComboBox = CurrentForm.DefaultView
        .ShowNavigationButtonsCheckBox = CurrentForm.NavigationButtons
        .NavigationCaptionTextBox = CurrentForm.NavigationCaption
        .RecordSelectorsCheckBox = CurrentForm.RecordSelectors
        .ModalCheckBox = CurrentForm.Modal
        .PopUpCheckBox = CurrentForm.PopUp
        .AutoCenterCheckBox = CurrentForm.AutoCenter
        .AllowAdditionsCheckBox = CurrentForm.AllowAdditions
        .AllowDeletionsCheckBox = CurrentForm.AllowDeletions
        .AllowEditsCheckBox = CurrentForm.AllowEdits
        .AllowFiltersCheckBox = CurrentForm.AllowFilters
    End With
    dialogForm.Visible = True
    dialogForm.SetFocus
    Do
        DoEvents
        If IsLoaded(stringDialogFormName) Then
            If dialogForm.Visible = False Then
                Exit Do
            End If
        End If
    Loop
    If Not IsLoaded(stringDialogFormName) Then Exit Function
    With dialogForm
         CurrentForm.Caption = .CaptionTextBox
         If .DefaultViewComboBox = 2 Then ' 2 is a datasheet
            CurrentForm.AllowDatasheetView = True
        End If
        If .DefaultViewComboBox = 0 Then ' 0 is a single form
            CurrentForm.AllowFormView = True
        End If
         CurrentForm.DefaultView = .DefaultViewComboBox
         CurrentForm.NavigationButtons = .ShowNavigationButtonsCheckBox
         CurrentForm.NavigationCaption = .NavigationCaptionTextBox
         CurrentForm.RecordSelectors = .RecordSelectorsCheckBox
         CurrentForm.Modal = .ModalCheckBox
         CurrentForm.PopUp = .PopUpCheckBox
         CurrentForm.AutoCenter = .AutoCenterCheckBox
         CurrentForm.AllowAdditions = .AllowAdditionsCheckBox
         CurrentForm.AllowDeletions = .AllowDeletionsCheckBox
         CurrentForm.AllowEdits = .AllowEditsCheckBox
         CurrentForm.AllowFilters = .AllowFiltersCheckBox
         DoCmd.Save acForm, CurrentForm.Name
    End With
    
        

ExitHere:
   Exit Function

SetFormProperties_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SetFormProperties of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

    
End Function

Public Sub CreateSelectQuery(stringObjectName As String)
    Dim stringQueryName As String
    Dim stringSQLText As String
    On Error GoTo CreateSelectQuery_Error
    If Len(stringObjectName & "") = 0 Then Exit Sub
    stringSQLText = " Select [" & stringObjectName & "].* From [" & stringObjectName & "];"
    Dim currentDatabase As DAO.database
    Set currentDatabase = CurrentDb
    Dim queryDefinition As DAO.QueryDef
    DoCmd.Close acQuery, "temp", acSaveYes
    On Error Resume Next
'    DoCmd.DeleteObject acQuery, "temp"
'    DoCmd.DeleteObject acQuery, "temp"
    On Error GoTo CreateSelectQuery_Error
    stringQueryName = "Temp" & Format(Now(), "yyyy-mmm-dd-hh-nn-ss")
    Set queryDefinition = currentDatabase.CreateQueryDef(stringQueryName, stringSQLText)
    DoCmd.OpenQuery stringQueryName, acViewDesign, acEdit


ExitHere:
   Exit Sub

CreateSelectQuery_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure CreateSelectQuery of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

    
End Sub



Public Function GetUsernameAndPassword(ByRef stringUsername As String, ByRef stringPassword As String) As Boolean
        Dim stringFormName As String: stringFormName = "Login"
        GetUsernameAndPassword = False
        If Not IsLoaded(stringFormName) Then
            DoCmd.Close acForm, stringFormName, acSaveYes
            DoCmd.OpenForm stringFormName, acNormal, , , acFormEdit, acDialog
        End If
        If Not IsLoaded(stringFormName) Then Exit Function
        Dim AccessForm As Access.Form
        Set AccessForm = Forms(stringFormName)
        stringUsername = AccessForm![UsernameTextBox]
        stringPassword = AccessForm![PasswordTextBox]
        GetUsernameAndPassword = True
        
End Function


Public Function LoadReportFromText()
    Dim stringFilename  As String
    
    On Error GoTo LoadReportFromText_Error
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "Text File", "*.txt"
        .Filters.Add "All Files", "*.*"
        .InitialView = msoFileDialogViewList
        .Title = "Select text file"
        If .Show Then
            stringFilename = .SelectedItems(1)
        End If
    End With
    If Len(stringFilename) = 0 Then Exit Function
    Dim stringReportName As String
    stringReportName = Mid(stringFilename, InStrRev(stringFilename, "\") + 1)
    stringReportName = Left(stringReportName, Len(stringReportName) - 4)
    Application.LoadFromText acReport, stringReportName, stringFilename

ExitHere:
   Exit Function

LoadReportFromText_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure LoadReportFromText of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Public Function SaveReportAsText()
    Dim stringReportName As String
    On Error GoTo SaveReportAsText_Error
    stringReportName = GetObjectName(acReport)
    Dim stringFolderName  As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Filters.Clear
        .InitialView = msoFileDialogViewList
        .Title = "Select the Folder where the report text file should be saved"
        If .Show Then
            stringFolderName = .SelectedItems(1)
        End If
    End With
    
    Application.SaveAsText acReport, stringReportName, stringFolderName & "\" & stringReportName & ".txt"

ExitHere:
   Exit Function

SaveReportAsText_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure SaveReportAsText of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Public Function HideAllButCurrent(blnShow As Boolean)
    Dim object As AccessObject, CurrentProject As CurrentProject
    On Error GoTo HideAllButCurrent_Error
    Dim blnActiveForm As Boolean
    Dim blnActiveReport As Boolean
    
    On Error Resume Next
    If Len(Screen.ActiveForm.Name) > 0 Then
        If Error.Number = 0 Then
            blnActiveForm = True
        Else
            blnActiveForm = False
        End If
    End If
    Error.Clear
    If Len(Screen.ActiveReport.Name) > 0 Then
        If Error.Number = 0 Then
            blnActiveReport = True
        Else
            blnActiveReport = False
        End If
    End If
    Dim AccessForm As Access.Form
    Set CurrentProject = Application.CurrentProject
    For Each object In CurrentProject.AllForms
        If object.IsLoaded Then
            If blnActiveForm Then
                If object.Name <> Screen.ActiveForm.Name Then
                    Set AccessForm = Forms(object.Name)
                    AccessForm.Visible = blnShow
                End If
            Else
                Set AccessForm = Forms(object.Name)
                AccessForm.Visible = blnShow
            End If
        End If
    Next object
    For Each object In CurrentProject.AllReports
        If object.IsLoaded Then
            If blnActiveReport Then
                Dim accessreport As Access.Report
                If object.Name <> Screen.ActiveReport.Name Then
                    Set accessreport = Reports(object.Name)
                    accessreport.Visible = blnShow
                End If
            Else
                Set accessreport = Reports(object.Name)
                accessreport.Visible = blnShow
            End If
        End If
    Next object
    Dim CodeProject As CodeProject
    Set CodeProject = Application.CodeProject
    For Each object In CodeProject.AllForms
        If object.IsLoaded Then
            If blnActiveForm Then
                If object.Name <> Screen.ActiveForm.Name Then
                    Set AccessForm = Forms(object.Name)
                    AccessForm.Visible = blnShow
                End If
            Else
                Set AccessForm = Forms(object.Name)
                AccessForm.Visible = blnShow
            End If
        End If
    Next object

ExitHere:
   Exit Function

HideAllButCurrent_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure HideAllButCurrent of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function

Public Function GetTableDefinition(currentDatabase As DAO.database, stringTableName As String) As DAO.tableDef
    Dim stringDatabase As String
    Dim blnAccessDatabase As Boolean
    On Error GoTo GetTableDefinition_Error
    Dim tableDefinition As DAO.tableDef
    
    Set tableDefinition = currentDatabase.TableDefs(stringTableName)
    Dim externalDatabase As DAO.database
    If InStr(tableDefinition.Connect, "DATABASE=") > 0 And InStr(tableDefinition.Connect, "ODBC") = 0 Then
        blnAccessDatabase = True
        stringDatabase = Mid(tableDefinition.Connect, InStr(tableDefinition.Connect, "DATABASE=") + 9)
        Set externalDatabase = DBEngine(0).OpenDatabase(stringDatabase)
        Set tableDefinition = externalDatabase.TableDefs(tableDefinition.SourceTableName)
        Set currentDatabase = externalDatabase
    Else
        blnAccessDatabase = False
        stringDatabase = ""
    End If
    Set GetTableDefinition = tableDefinition

ExitHere:
   Exit Function

GetTableDefinition_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure GetTableDefinition of Module basAddInProcedures" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Function