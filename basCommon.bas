Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All rights reserved.

' Required by:
'   basFileOpen
'   basFontHandling
'   basObjList
'   CommonDlg
'   ScreenInfo
'   ShellBrowse
'    VersionInfo

' Common routines needed by many of the procedures
' in this project.

' Registry Errors, used in several modules.

Public Enum adhRegErrors
    adhcAccErrSuccess = 0
    adhcAccErrUnknown = -1
    adhcAccErrRegKeyNotFound = -201
    adhcAccErrRegValueNotFound = -202
    adhcAccErrRegCantSetValue = -203
    adhcAccErrRegSubKeyNotFound = -204
    adhcAccErrRegTypeNotSupported = -205
    adhcAccErrRegCantCreateKey = -206
    adhcAccErrRegBufferTooSmall = -207
    adhcAccErrRegCantDeleteValue = -208
End Enum

Private Type Ampersand
    strValue As String
    intPosnOfAmpersand As Integer
End Type

' ================
' GetWindowRect Info
' ================
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Public Declare Function GetWindowRect _
' Lib "User32" _
' (ByVal hwnd As Long, lpRect As RECT) As Long
'
'
'Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
'   (ByVal hwndCaller As Long, _
'   ByVal pszFile As String, _
'   ByVal uCommand As Long, _
'   dwData As Any) As Long

' Indicate that a parameter for QuickSort is missing.
Private Const dhcMissing = -2

Public Function adhFnPtrToLong(lngAddress As Long) As Long
    
    ' Given a function pointer as a Long, return a Long.
    ' Sure looks like this function isn't doing anything,
    ' and in reality, it's not.
    
    ' Call this function like this:
    '
    ' lngPointer = adhFnPtrToLong(AddressOf SomeFunction)
    
    ' and it returns the address you've sent it
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.

    ' In:
    '   lngAddress:
    '       address of a public procedure, passed using
    '       the AddressOf modifier.
    ' Out:
    '   Return value:
    '       The input address, cast as a Long.


On Error GoTo HandleErr

    adhFnPtrToLong = lngAddress

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.adhFnPtrToLong"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Function adhTrimNull(strVal As String) As String
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.
    
    ' Trim the end of a string, stopping at the first
    ' null character.
    
    Dim intPos As Integer

On Error GoTo HandleErr

    intPos = InStr(1, strVal, vbNullChar)
    Select Case intPos
        Case Is > 1
            adhTrimNull = Left$(strVal, intPos - 1)
        Case 0
            adhTrimNull = strVal
        Case 1
            adhTrimNull = vbNullString
    End Select

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.adhTrimNull"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Sub adhQuickSort(varArray As Variant, _
 Optional intLeft As Integer = dhcMissing, _
 Optional intRight As Integer = dhcMissing)

    ' Quicksort for simple data types.

    ' From Access 2002 Desktop Developer's Handbook
    ' by Litwin, Getz, and Gunderloy. (Sybex)
    ' Copyright 2001. All rights reserved.
    
    ' Originally from "VBA Developer's Handbook"
    ' by Ken Getz and Mike Gilbert
    ' Copyright 1997; Sybex, Inc. All rights reserved.
    
    ' Entry point for sorting the array.
    
    ' This technique uses the recursive Quicksort
    ' algorithm to perform its sort.
    
    ' In:
    '   varArray:
    '       A variant pointing to an array to be sorted.
    '       This had better actually be an array, or the
    '       code will fail, miserably. You could add
    '       a test for this:
    '       If Not IsArray(varArray) Then Exit Sub
    '       but hey, that would slow this down, and it's
    '       only YOU calling this procedure.
    '       Make sure it's an array. It's your problem.
    '   intLeft:
    '   intRight:
    '       Lower and upper bounds of the array to be sorted.
    '       If you don't supply these values (and normally, you won't)
    '       the code uses the LBound and UBound functions
    '       to get the information. In recursive calls
    '       to the sort, the caller will pass this information in.
    '       To allow for passing integers around (instead of
    '       larger, slower variants), the code uses -2 to indicate
    '       that you've not passed a value. This means that you won't
    '       be able to use this mechanism to sort arrays with negative
    '       indexes, unless you modify this code.
    ' Out:
    '       The data in varArray will be sorted.
    
    Dim i As Integer
    Dim j As Integer
    Dim varTestVal As Variant
    Dim intMid As Integer


On Error GoTo HandleErr

    If intLeft = dhcMissing Then intLeft = LBound(varArray)
    If intRight = dhcMissing Then intRight = UBound(varArray)
   
    If intLeft < intRight Then
        intMid = (intLeft + intRight) \ 2
        varTestVal = UCase(varArray(intMid))
        i = intLeft
        j = intRight
        Do
            Do While UCase(varArray(i)) < varTestVal
                i = i + 1
            Loop
            Do While UCase(varArray(j)) > varTestVal
                j = j - 1
            Loop
            If i <= j Then
                Call SwapElements(varArray, i, j)
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= intMid Then
            Call adhQuickSort(varArray, intLeft, j)
            Call adhQuickSort(varArray, i, intRight)
        Else
            Call adhQuickSort(varArray, i, intRight)
            Call adhQuickSort(varArray, intLeft, j)
        End If
    End If

ExitHere:
  Exit Sub

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.adhQuickSort"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Sub

Private Sub SwapElements(varItems As Variant, intItem1 As Integer, intItem2 As Integer)
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All rights reserved.
    
    Dim varTemp As Variant


On Error GoTo HandleErr

    varTemp = varItems(intItem2)
    varItems(intItem2) = varItems(intItem1)
    varItems(intItem1) = varTemp

ExitHere:
  Exit Sub

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.SwapElements"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Sub

Public Function RunDocumentorAndDisplayReports() As Boolean

On Error GoTo HandleErr

' TO DO: Turn normal error handler on when this condition is finished.
On Error Resume Next

Dim dbAddin As DAO.database
Dim dbRemote As DAO.database
Dim rStd As DAO.Recordset
Dim rsQD As DAO.Recordset
Dim rsQS As DAO.Recordset
Dim strSQL As String
Dim DocName As String

    Set dbAddin = codeDB
    Set dbRemote = CurrentDb
    
    ' Delete old entries from all three tables
    If DCount("[TblName]", "a_TableMetaData") <> 0 Then
        strSQL = "DELETE * FROM a_TableMetaData"
        dbAddin.Execute strSQL
    End If
    
    If DCount("[QryName]", "a_QueryMetaData") <> 0 Then
        strSQL = "DELETE * FROM a_QueryMetaData"
        dbAddin.Execute strSQL
    End If
    
    If DCount("[QryName]", "a_QuerySQL") <> 0 Then
        strSQL = "DELETE * FROM a_QuerySQL"
        dbAddin.Execute strSQL
    End If
    
    '  Get table/field information
    Set rStd = dbAddin.OpenRecordset("a_TableMetaData")
    Dim QName As String
    Dim SQLCode As String
    
    Dim T As Integer
    Dim f As Integer
    Dim Q As Integer
    
    Dim TName As String
    Dim fname As String
    Dim FType As Variant
    Dim FSize As Integer
    
    For T = 0 To dbRemote.TableDefs.Count - 1
        If Left(dbRemote.TableDefs(T).Name, 4) <> "Msys" Then
            TName = dbRemote.TableDefs(T).Name
            For f = 0 To dbRemote.TableDefs(T).Fields.Count - 1
                fname = dbRemote.TableDefs(T).Fields(f).Name
                FType = dbRemote.TableDefs(T).Fields(f).Type
                Select Case FType
                    Case DB_BOOLEAN:  FType = "Yes/No"
                    Case DB_BYTE:  FType = "Number (Byte)"
                    Case DB_INTEGER:  FType = "Number (Integer)"
                    Case DB_LONG
                        If (dbRemote.TableDefs(T).Fields(f).Attributes And dbAutoIncrField) Then
                            FType = "Counter"
                        Else
                            FType = "Number (Long Integer)"
                        End If
                    Case DB_CURRENCY:  FType = "Currency"
                    Case DB_SINGLE:  FType = "Number (Single)"
                    Case DB_DOUBLE:  FType = "Number (Double)"
                    Case DB_DATE:  FType = "Date/Time"
                    Case DB_TEXT:  FType = "Text"
                    Case DB_LONGBINARY:  FType = "OLE Object"
                    Case DB_MEMO:  FType = "Memo"
                    'Case dbGUID:  FType = "Replication ID"
                    Case Else:  FType = "Unknown"
                End Select
        
                FSize = dbRemote.TableDefs(T).Fields(f).Size
    
                rStd.AddNew
                rStd!TblName = TName
                rStd!FldName = fname
                rStd!FldType = FType
                rStd!FldSize = FSize
                rStd.Update
            Next f
        End If
    Next T
    rStd.Close
    
    
    '  Get Query/SQL Information
    Set rsQS = dbAddin.OpenRecordset("a_QuerySQL")
    Set rsQD = dbAddin.OpenRecordset("a_QueryMetaData")
    f = 0
    
    For Q = 0 To dbRemote.QueryDefs.Count - 1
        QName = dbRemote.QueryDefs(Q).Name
        SQLCode = dbRemote.QueryDefs(Q).SQL
    
        For f = 0 To dbRemote.QueryDefs(Q).Fields.Count - 1
            rsQD.AddNew
            fname = dbRemote.QueryDefs(Q).Fields(f).Name
            TName = dbRemote.QueryDefs(Q).Fields(f).SourceTable
    
            If Len(TName) = 0 Then
                TName = "Calculated Field"
            End If
    
            rsQD!QryName = QName
            rsQD!TblName = TName
            rsQD!FldName = fname
            rsQD.Update
        Next f
    
        rsQS.AddNew
        rsQS!QryName = QName
        rsQS!SQLCode = SQLCode
        rsQS.Update
    Next Q
    
    rsQD.Close
    rsQS.Close
    dbAddin.Close
    dbRemote.Close

    DocName = "rptDocumentation_QUERIES"
    DoCmd.OpenReport DocName, A_PREVIEW
    DoCmd.MoveSize 200, 200, 8000, 5000

    DocName = "rptDocumentation_TABLES"
    DoCmd.OpenReport DocName, A_PREVIEW
    DoCmd.MoveSize 500, 500, 8000, 5000
    
    DocName = "rptDocumentation_Other"
    DoCmd.OpenReport DocName, A_PREVIEW
    DoCmd.MoveSize 800, 800, 8000, 5000
    

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.RunDocumentorAndDisplayReports"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Function GetDirectory()
 Dim strDir As String
On Error GoTo HandleErr
    Dim stringFilename  As String
                   With Application.FileDialog(msoFileDialogFolderPicker)
                       .InitialView = msoFileDialogViewList
                       .Title = "Select a Folder"
                       If .Show Then
                           strDir = .SelectedItems(1)
                       End If
                   End With
           
'    strDir = GetFileNameOfficeDialog(True, True, False, True, True, 0, "", 1, "", "Choose a Directory", "Select", "C:\", True)
    If Len(strDir) > 0 Then
        Screen.activeControl = strDir
        'Screen.ActiveDatasheet.ActiveControl = strDir
    End If
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case # '
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.GetDirectory"    'ErrorHandler:$$N=basCommon.GetDirectory
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function
Public Function getfile()
 Dim strDir As String
On Error GoTo HandleErr
            Dim stringFilename  As String
            With Application.FileDialog(msoFileDialogOpen)
                .Filters.Clear
                .Filters.Add "Any File", "*.*"
                .InitialView = msoFileDialogViewList
                .Title = "Select a file"
                If .Show Then
                    stringFilename = .SelectedItems(1)
                End If
            End With

    strDir = stringFilename
    If Len(strDir) > 0 Then
        Screen.activeControl = strDir
        'Screen.ActiveDatasheet.ActiveControl = strDir
    End If
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case # '
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.GetDirectory"    'ErrorHandler:$$N=basCommon.GetDirectory
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function



Public Function FieldType(intType As Integer) As String


On Error GoTo HandleErr

    Select Case intType
        Case dbBoolean
            FieldType = "dbBoolean"
        Case dbByte
            FieldType = "dbByte"
        Case dbInteger
            FieldType = "dbInteger"
        Case dbLong
            FieldType = "dbLong"
        Case dbCurrency
            FieldType = "dbCurrency"
        Case dbSingle
            FieldType = "dbSingle"
        Case dbDouble
            FieldType = "dbDouble"
        Case dbDate
            FieldType = "dbDate"
        Case dbText
            FieldType = "dbText"
        Case dbLongBinary
            FieldType = "dbLongBinary"
        Case dbMemo
            FieldType = "dbMemo"
        Case dbGUID
            FieldType = "dbGUID"
        Case dbAttachment
            FieldType = "dbAttachment"
        Case dbBinary
            FieldType = "dbBinary"
        Case dbDecimal
            FieldType = "dbDecimal"
        Case Else
            FieldType = "Undefined"
    End Select


ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.FieldType"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Function FormattedMsgBox( _
 Prompt As String, _
 Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
 Optional Title As String = vbNullString, _
 Optional HelpFile As Variant, _
 Optional Context As Variant) _
 As VbMsgBoxResult
    

On Error GoTo HandleErr

    If IsMissing(HelpFile) Or IsMissing(Context) Then
        FormattedMsgBox = Eval("MsgBox(""" & Prompt & _
         """, " & Buttons & ", """ & Title & """)")
    Else
        FormattedMsgBox = Eval("MsgBox(""" & Prompt & _
         """, " & Buttons & ", """ & Title & """, """ & _
         HelpFile & """, " & Context & ")")
    End If

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.FormattedMsgBox"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function


Public Function AllProcs(strModuleName As String) As Variant
    Dim module As module
    Dim lngCount As Long, lngCountDecl As Long, lngI As Long
    Dim strProcName As String, astrProcNames() As String
    Dim intI As Integer, strMsg As String
    Dim lngR As Long
    On Error GoTo HandleErr
    ' Open specified Module object.
    DoCmd.OpenModule strModuleName
    ' Return reference to Module object.
    Set module = Modules(strModuleName)
    ' Count lines in module.
    lngCount = module.CountOfLines
    ' Count lines in Declaration section in module.
    lngCountDecl = module.CountOfDeclarationLines
    ' Determine name of first procedure.
    strProcName = module.ProcOfLine(lngCountDecl + 1, lngR)
    ' Initialize counter variable.
    intI = 0
    ' Redimension array.
    ReDim Preserve astrProcNames(intI)
    ' Store name of first procedure in array.
    astrProcNames(intI) = strProcName
    ' Determine procedure name for each line after declarations.
    For lngI = lngCountDecl + 1 To lngCount
        ' Compare procedure name with ProcOfLine property value.
        If strProcName <> module.ProcOfLine(lngI, lngR) Then
            ' Increment counter.
            intI = intI + 1
            strProcName = module.ProcOfLine(lngI, lngR)
            ReDim Preserve astrProcNames(intI)
            ' Assign unique procedure names to array.
            astrProcNames(intI) = strProcName
        End If
    Next lngI
    
    Call adh_accSortStringArray(astrProcNames())
    For intI = 0 To UBound(astrProcNames)
        strMsg = strMsg & astrProcNames(intI) & ";"
    Next intI
    AllProcs = strMsg
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 2017 'VBA Code protected - please supply a password
        MsgBox "The VBA Code is protected - please unlock by supplying a password first", vbInformation, "Code Protected"
        DoCmd.RunCommand acCmdVisualBasicEditor
        Resume ExitHere
    Case 2517  'Access can't find the procedure
        AllProcs = ""
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.AllProcs" 'ErrorHandler:$$N=basCommon.AllProcs
        Resume ExitHere
    End Select
    Resume 'Debug Only
End Function

Public Function RenameControlsToRVBAConvention(strObject As String, acObject As Access.AcObjectType) As Boolean
    Dim dbCode As DAO.database
    Dim rs As DAO.Recordset
    Dim rstR As DAO.Recordset
    Dim ctl As Access.Control
    Dim aobCurrent As Object
    Dim strSQLtext As String
    Dim strMsg As String, strHeading As String
    Dim intCounter As Integer, strControl As String
    Dim strNewName As String
    Dim strTemp As String, k As Integer
    RenameControlsToRVBAConvention = False
    On Error GoTo HandleErr
    Select Case acObject
    Case acForm
        strMsg = "Warning this process will add the tag suffix to all controls on " _
        & strObject & " Form." & vbCrLf & "If the controls have events or are referenced in code/queries, then they will no longer work!" & vbCrLf & "Proceed?"
    Case acReport
        strMsg = "Warning this process will add the tag suffix to all controls on " _
        & strObject & " Report." & vbCrLf _
        & "If the controls are referenced in code they will no longer work!" & vbCrLf & "Proceed?"
    End Select
    
'    Select Case MsgBox(strMsg, vbYesNo + vbExclamation + vbDefaultButton2, "Rename to RVBA Convention")
'    Case vbNo
'        Exit Function
'    End Select
    Select Case acObject
    Case acForm
        DoCmd.OpenForm strObject, acDesign
        Set aobCurrent = Forms(strObject)
    Case acReport
        DoCmd.OpenReport strObject, acDesign
        Set aobCurrent = Reports(strObject)
    End Select
    
    Set dbCode = codeDB
    Set rstR = dbCode.OpenRecordset("tmpCtlsRenamed")
    ZapTable rstR
    For Each ctl In aobCurrent.Controls
        strSQLtext = "SELECT tblRVBAAccessObjTags.*" & vbCrLf
        strSQLtext = strSQLtext & "           , tblRVBAAccessObjTags.[Type No]" & vbCrLf
        strSQLtext = strSQLtext & "        FROM tblRVBAAccessObjTags" & vbCrLf
        strSQLtext = strSQLtext & "       WHERE (((tblRVBAAccessObjTags.[Type No])=" & ctl.ControlType & "));"
        Set rs = dbCode.OpenRecordset(strSQLtext)
        If Not rs.EOF Then
            rs.MoveFirst
            If InStr(ctl.Name, rs![tag]) = 0 Then
                intCounter = intCounter + 1
                rstR.AddNew
                rstR![Old Name] = Left(ctl.Name, 255)
                strTemp = ""
                For k = 1 To Len(ctl.Name)
                    Select Case Mid(ctl.Name, k, 1)
                    Case " ", ".", ","
                    Case Else
                        strTemp = strTemp & Mid(ctl.Name, k, 1)
                    End Select
                Next k
                If Right(strTemp, 6) = "_Label" Then
                    strTemp = Left(strTemp, Len(strTemp) - 6)
                End If
                strNewName = strTemp
                rstR![New Name] = strNewName & rs![tag]
                rstR![Object Type] = rs![Object Type]
                rstR.Update
            End If
        End If
    Next
    Select Case intCounter
    Case 0
        strHeading = "All OK"
        strMsg = "There were no controls renamed"
    Case 1
        strHeading = "1 Control Found"
        strMsg = "There is 1 control to rename"
    Case Is > 1
        strHeading = "Some Controls Found"
        strMsg = "There are " & intCounter & " controls to rename"
    End Select
    If intCounter > 0 Then
        DoCmd.OpenForm "frmCtlsRenamed", acNormal, , , acFormEdit, acWindowNormal, strObject
    End If
'    MsgBox strMsg, vbInformation, strHeading
    RenameControlsToRVBAConvention = True
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case # '
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.RenameControlsToRVBAConvention"  'ErrorHandler:$$N=basCommon.RenameControlsToRVBAConvention
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function

Public Function CheckSpellingofCaptions(strObject As String, acObject As Access.AcObjectType _
    , Optional blnBatch As Boolean = False) As Boolean
    Dim strDB  As String 'Added 21/06/2004
    Dim dbCode As DAO.database
    Dim DBProj As DAO.database
    Dim td As DAO.tableDef
    Dim fld As DAO.field
    Dim rs As DAO.Recordset
    Dim ctl As Access.Control
    Dim aobCurrent As Object
    Dim strSQLtext As String
    Dim strMsg As String, strHeading As String
    Dim intCounter As Integer, strControl As String
    Dim strCheckText As String
    Dim strTemp As String, k As Integer
    Dim frm As Access.Form
    Dim rpt As Access.Report
    Dim lngKouter As Long
    Dim ra As Ampersand
    Dim db1 As DAO.database
    On Error GoTo HandleErr
    CheckSpellingofCaptions = False
    DoCmd.Close acForm, "frmMisspelledPropertiesWrapper", acSaveYes
    Select Case acObject
    Case acForm
        strMsg = "This process will check the spelling of all Captions Status Bar Text, Validation Text and Control Tool Tips on " _
        & strObject & " Form." & vbCrLf & "" & vbCrLf & "Proceed?"
    Case acTable
        strMsg = "This process will check the spelling of all Field Descriptions, Captions and Validation Text on " _
        & strObject & " Table." & vbCrLf & "" & vbCrLf & "Proceed?"
    Case acReport
        strMsg = "This process will check the spelling of all Captions on " _
        & strObject & " Report." & vbCrLf & "" & vbCrLf & "Proceed?"
    Case Else
        Exit Function
    End Select
    
    Select Case MsgBox(strMsg, vbYesNo + vbExclamation + vbDefaultButton2, "Check Spelling")
    Case vbNo
        Exit Function
    End Select
    Set dbCode = codeDB
    Set rs = dbCode.OpenRecordset("tmpMisspelledProperties")
    ZapTable rs
    Select Case acObject
    Case acForm
        DoCmd.OpenForm strObject, acDesign
        Set frm = Forms(strObject)
        Set aobCurrent = Forms(strObject)
        SysCmd acSysCmdInitMeter, "Building a list of properties...", aobCurrent.Controls.Count
        lngKouter = 0
        For Each ctl In aobCurrent.Controls
            lngKouter = lngKouter + 1
            If ctl.ControlType = acLabel Or ctl.ControlType = acToggleButton Or ctl.ControlType = acCommandButton Then
                If Len(ctl.Caption & "") > 1 Then
                    ra = RemoveAmpersand(ctl.Caption)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = ctl.Name
                    rs![ControlType] = ctl.ControlType
                    rs![PropertyType] = "Caption"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
                End If
            End If
            If ctl.ControlType = acLabel Or ctl.ControlType = acTextBox Or ctl.ControlType = acOptionGroup _
            Or ctl.ControlType = acToggleButton Or ctl.ControlType = acCheckBox Or ctl.ControlType = acComboBox _
             Or ctl.ControlType = acListBox Or ctl.ControlType = acCommandButton _
              Or ctl.ControlType = acImage Or ctl.ControlType = acBoundObjectFrame Then
                If Len(ctl.ControlTipText & "") > 1 Then
                    ra = RemoveAmpersand(ctl.ControlTipText)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = ctl.Name
                    rs![ControlType] = ctl.ControlType
                    rs![PropertyType] = "ControlTipText"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
                End If
            End If
            If ctl.ControlType = acTextBox Or ctl.ControlType = acOptionGroup Or ctl.ControlType = acToggleButton _
            Or ctl.ControlType = acCheckBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acListBox _
             Or ctl.ControlType = acCommandButton Or ctl.ControlType = acBoundObjectFrame _
             Or ctl.ControlType = acTabCtl Or ctl.ControlType = acSubform Then
                If Len(ctl.StatusBarText & "") > 1 Then
                    ra = RemoveAmpersand(ctl.StatusBarText)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = ctl.Name
                    rs![ControlType] = ctl.ControlType
                    rs![PropertyType] = "StatusBarText"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
                End If
            End If
            If ctl.ControlType = acTextBox Or ctl.ControlType = acOptionGroup _
            Or ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
                If Len(ctl.ValidationText & "") > 1 Then
                    ra = RemoveAmpersand(ctl.ValidationText)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = ctl.Name
                    rs![ControlType] = ctl.ControlType
                    rs![PropertyType] = "ValidationText"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
                End If
            End If
            SysCmd acSysCmdUpdateMeter, lngKouter
        Next
        SysCmd acSysCmdRemoveMeter
        'Form Caption
        ra = RemoveAmpersand(frm.Caption)
        rs.AddNew
        rs![CheckTextBefore] = ra.strValue
        rs![CheckTextAfter] = ra.strValue
        rs![ControlName] = frm.Name
        rs![ControlType] = 97
        rs![PropertyType] = "FormCaption"
        rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
        rs.Update
    Case acReport
        DoCmd.OpenReport strObject, acDesign
        Set rpt = Reports(strObject)
        Set aobCurrent = Reports(strObject)
        SysCmd acSysCmdInitMeter, "Building a list of properties...", aobCurrent.Controls.Count
        lngKouter = 0
        For Each ctl In aobCurrent.Controls
            lngKouter = lngKouter + 1
            If ctl.ControlType = acLabel Or ctl.ControlType = acToggleButton Or ctl.ControlType = acCommandButton Then
                If Len(ctl.Caption & "") > 1 Then
                    ra = RemoveAmpersand(ctl.Caption)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = ctl.Name
                    rs![ControlType] = ctl.ControlType
                    rs![PropertyType] = "Caption"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
                End If
            End If
            SysCmd acSysCmdUpdateMeter, lngKouter
        Next
        SysCmd acSysCmdRemoveMeter
        'Report Caption
        ra = RemoveAmpersand(rpt.Caption)
        rs.AddNew
        rs![CheckTextBefore] = ra.strValue
        rs![CheckTextAfter] = ra.strValue
        rs![ControlName] = rpt.Name
        rs![ControlType] = 98
        rs![PropertyType] = "ReportCaption"
        rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
        rs.Update
    Case acTable
        Set DBProj = CurrentDb
        Set td = DBProj.TableDefs(strObject)
        If td.Attributes = dbAttachedTable Then
            strDB = Mid(td.Connect, 11)
            Select Case MsgBox("This table is a linked table.  " _
            & "Note spelling will be checked in the back end database where the table resides." _
                 & vbCrLf & "Not the Linked table." _
                 & vbCrLf & vbCrLf & "Database: " & strDB _
            , vbOKCancel + vbQuestion + vbDefaultButton2, "Table is Linked...")
            Case vbCancel
                Exit Function
            End Select
            Set db1 = DBEngine.Workspaces(0).OpenDatabase(strDB)
        ElseIf td.Attributes = dbAttachedODBC + dbAttachSavePWD Or td.Attributes = dbAttachedODBC Then
            MsgBox "This table is not a native access table", vbExclamation, "Aborting... ODBC Linked"
            Exit Function
        Else
            Set db1 = CurrentDb
        End If
        SysCmd acSysCmdInitMeter, "Walking fields...", td.Fields.Count
        For Each fld In td.Fields
            intCounter = intCounter + 1
            strTemp = CStr(GetProperty_DAO(strObject, fld.Name, "Description", db1))
            If Len(strTemp) > 0 Then
                    ra = RemoveAmpersand(strTemp)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = fld.Name
                    rs![ControlType] = 99
                    rs![PropertyType] = "Field Description"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
            End If
            strTemp = CStr(GetProperty_DAO(strObject, fld.Name, "Caption", db1))
            If Len(strTemp) > 0 Then
                    ra = RemoveAmpersand(strTemp)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = fld.Name
                    rs![ControlType] = 99
                    rs![PropertyType] = "Field Caption"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
            End If
            strTemp = fld.ValidationText
            If Len(strTemp) > 0 Then
                    ra = RemoveAmpersand(strTemp)
                    rs.AddNew
                    rs![CheckTextBefore] = ra.strValue
                    rs![CheckTextAfter] = ra.strValue
                    rs![ControlName] = fld.Name
                    rs![ControlType] = 99
                    rs![PropertyType] = "ValidationText"
                    rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
                    rs.Update
            End If
            SysCmd acSysCmdUpdateMeter, intCounter
        Next
        SysCmd acSysCmdRemoveMeter
        'Table Validation Text
        strTemp = td.ValidationText
        If Len(strTemp) > 0 Then
            ra = RemoveAmpersand(strTemp)
            rs.AddNew
            rs![CheckTextBefore] = ra.strValue
            rs![CheckTextAfter] = ra.strValue
            rs![ControlName] = td.Name
            rs![ControlType] = 96
            rs![PropertyType] = "TableValidationText"
            rs![PosnOfAmpersand] = ra.intPosnOfAmpersand
            rs.Update
        End If
    End Select
    
    If blnBatch Then 'Open form in Dialog mode
        CheckSpellingofCaptions = True
        DoCmd.OpenForm "frmMisspelledPropertiesWrapper", acNormal, , , acFormEdit, acDialog, strObject & "|" & acObject
    Else
        DoCmd.OpenForm "frmMisspelledPropertiesWrapper", acNormal, , , acFormEdit, acWindowNormal, strObject & "|" & acObject
        CheckSpellingofCaptions = True
    End If
    

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case # '
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.CheckSpellingofCaptions" 'ErrorHandler:$$N=basCommon.CheckSpellingofCaptions
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function


Function GetProperty_DAO(ByVal MyTableName As String, _
  ByVal MyFieldName As String, ByVal MyProperty As String, Optional db As DAO.database)
    Dim td As DAO.tableDef
    Dim fld As DAO.field
    
    On Error GoTo Err_GetFieldDescription
    If db Is Nothing Then
        Set db = CurrentDb
    End If
    Set td = db.TableDefs(MyTableName)
    Set fld = td.Fields(MyFieldName)
    GetProperty_DAO = fld.Properties(MyProperty)
    
Bye_GetFieldDescription:
    Exit Function
    
Err_GetFieldDescription:
    GetProperty_DAO = ""
    
    Resume Bye_GetFieldDescription

End Function



Public Function CheckAccelerators(strObject As String, acObject As Access.AcObjectType _
    , Optional blnBatch As Boolean = False) As Boolean
    
    Dim f As Access.Form
    Dim sfm As Access.Form
    Dim ctl As Access.Control
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim ctl2 As Access.Control
    Dim ctl3 As Access.Control
    Dim ctlsub As Access.Control
    Dim strTemp As String
    On Error GoTo HandleErr
    Set db = codeDB
    Set rs = db.OpenRecordset("tmpAccelerators")
    ZapTable rs
    CheckAccelerators = False
    If acObject <> acForm Then Exit Function
    DoCmd.OpenForm strObject, acDesign
    Set f = Forms(strObject)
    For Each ctl In f.Controls
        strTemp = ""
        Select Case ctl.Section
        Case acHeader
            strTemp = "Header"
        Case acDetail
            strTemp = "Detail"
        Case acFooter
            strTemp = "Footer"
        Case acPageHeader 'Accelerators are of no use here
            strTemp = "PageHeader"
        Case acPageFooter 'Accelerators are of no use here
            strTemp = "PageFooter"
        End Select
        Select Case ctl.ControlType
        Case acLabel
            If ctl.Parent.Name <> strObject And Len(ctl.Caption) < 256 Then 'Not interested in single lables
                rs.AddNew
                rs![ControlName] = ctl.Parent.Name
                Select Case ctl.Parent.ControlType
                Case acTextBox, acComboBox, acCheckBox, acBoundObjectFrame, acListBox, acOptionButton
                    rs![Locked] = ctl.Parent.Locked
                    rs![Enabled] = ctl.Parent.Enabled
                    If Not rs![Enabled] Then
                        rs![Exclude] = True
                    End If
                End Select
                Set ctl2 = f.Controls(rs![ControlName])
                If ctl2.Parent.Name <> strObject Then 'Label is on a tab page or part or a group
                    Set ctl3 = ctl2.Parent
                    If ctl3.ControlType = acPage Then
                        rs![TabPageName] = ctl3.Name
                    End If
                End If
                rs![Container] = strTemp
                rs![LabelName] = ctl.Name
                If InStr(ctl.Caption, "&") > 0 And InStr(ctl.Caption, "&") < Len(ctl.Caption) Then
                    rs![Accelerator] = Mid(ctl.Caption, InStr(ctl.Caption, "&") + 1, 1)
                End If
                rs![Caption] = ctl.Caption
                rs![FormName] = strObject
                rs.Update
            End If
        Case acToggleButton
            rs.AddNew
            rs![ControlName] = ctl.Parent.Name
            If rs![ControlName] <> strObject Then
                Set ctl2 = f.Controls(rs![ControlName])
                If ctl2.ControlType = acPage Then
                    rs![TabPageName] = ctl2.Name
                End If
            Else
            End If
            rs![Container] = strTemp
            rs![LabelName] = ctl.Name
            If InStr(ctl.Caption, "&") > 0 And InStr(ctl.Caption, "&") < Len(ctl.Caption) Then
                rs![Accelerator] = Mid(ctl.Caption, InStr(ctl.Caption, "&") + 1, 1)
            End If
            rs![Locked] = ctl.Locked
            rs![Enabled] = ctl.Enabled
            rs![Caption] = ctl.Caption
            rs![FormName] = strObject
            rs.Update
        Case acCommandButton
            rs.AddNew
            rs![Container] = strTemp
            rs![LabelName] = ctl.Name
            If InStr(ctl.Caption, "&") > 0 And InStr(ctl.Caption, "&") < Len(ctl.Caption) Then
                rs![Accelerator] = Mid(ctl.Caption, InStr(ctl.Caption, "&") + 1, 1)
            End If
            rs![Caption] = ctl.Caption
            rs![Enabled] = ctl.Enabled
            rs![FormName] = strObject
            rs.Update
        Case acPage
            rs.AddNew
            rs![ControlName] = ctl.Parent.Name
            rs![Container] = strTemp
            rs![LabelName] = ctl.Name
            If InStr(ctl.Caption, "&") > 0 And InStr(ctl.Caption, "&") < Len(ctl.Caption) Then
                rs![Accelerator] = Mid(ctl.Caption, InStr(ctl.Caption, "&") + 1, 1)
            End If
            rs![Caption] = ctl.Caption
            rs![Enabled] = ctl.Enabled
            rs![FormName] = strObject
            rs.Update
        Case acSubform
                ' If form has sub form add these as well
                    If Len(Forms(strObject)(ctl.Name).SourceObject & "") > 0 Then
                        Set sfm = Forms(strObject)(ctl.Name).Form
                        
                        For Each ctlsub In sfm.Controls
                            Select Case ctlsub.ControlType
                            Case acLabel
                                If ctlsub.Parent.Name <> sfm.Name Then  'Not interested in single lables
                                    rs.AddNew
                                    rs![ControlName] = ctlsub.Parent.Name
                                    Set ctl2 = sfm.Controls(rs![ControlName])
                                    If ctl.Parent.Name <> strObject Then ' check it is not a form
                                        If ctl.Parent.ControlType = acPage Then
                                            rs![TabPageName] = ctl.Parent.Name
                                        End If
                                    End If
                                    strTemp = ""
                                    Select Case ctl.Section
                                    Case acHeader
                                        strTemp = "Header"
                                    Case acDetail
                                        strTemp = "Detail"
                                    Case acFooter
                                        strTemp = "Footer"
                                    End Select
                                    rs![Container] = strTemp
                                    rs![LabelName] = ctlsub.Name
                                    If InStr(ctlsub.Caption, "&") > 0 And InStr(ctlsub.Caption, "&") < Len(ctlsub.Caption) Then
                                        rs![Accelerator] = Mid(ctlsub.Caption, InStr(ctlsub.Caption, "&") + 1, 1)
                                    End If
                                    rs![Caption] = ctlsub.Caption
                                    rs![FormName] = sfm.Name
                                    rs![subformControl] = ctl.Name
                                    rs.Update
                                End If
                            End Select
                        Next ctlsub
                    End If
        End Select
    Next ctl
    rs.Close
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qryAcceleratorsMarkDupes", acViewNormal
    DoCmd.SetWarnings True
    DoCmd.OpenForm "frmAccelerators", acNormal, , "[Duplicated] = " & True, acFormEdit, acWindowNormal
    Forms![frmAccelerators]![txtFormName] = strObject
    Forms![frmAccelerators].Caption = "Check Accelerator Keys for " & strObject
    CheckAccelerators = True

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case # '
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.CheckAccelerators"   'ErrorHandler:$$N=basCommon.CheckAccelerators
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function


Public Function adh_accSortStringArray _
 (astrObjects() As String) As Long
 
    ' From Access 2002 Desktop Developer's Handbook
    ' by Litwin, Getz, and Gunderloy. (Sybex)
    ' Copyright 2001. All rights reserved.
    
    ' Sort an array of strings. Just a wrapper
    ' around the common adhQuickSort procedure.
    
    ' In:
    '   astrObjects():
    '       An array of strings. There should be
    '       no blank entries, or those will
    '       sort to the top of the array.
    ' Out:
    '   astrObjects():
    '       Now sorted, alphabetically.
    '   Return Value:
    '       adhcAccErrSuccess or adhcAccErrUnknown
    
    Dim lngRetVal As Long
    
    On Error GoTo HandleErrors
    Call adhQuickSort(astrObjects)
    lngRetVal = adhcAccErrSuccess
    
ExitHere:
    adh_accSortStringArray = lngRetVal
    Exit Function
    
HandleErrors:
    ' Just return an error value
    ' if any error occurs. You may
    ' want to investigate, however...
    lngRetVal = adhcAccErrUnknown
    Resume ExitHere
End Function

Public Function StartProgressBars(lngMainMax As Long, lngSubMax As Long, _
Optional strMainCaption As String = "Main Process", _
Optional strSubCaption As String = "Sub Process")
    Dim f As Access.Form

On Error GoTo HandleErr

    DoCmd.OpenForm "frmProgressBar", acNormal, , , acFormEdit, acWindowNormal
    Set f = Forms![frmProgressBar]
    f.ProgressBarMain.Min = 0
    f.ProgressBarMain.Max = lngMainMax
    f.ProgressBarSub.Min = 0
    f.lblMain.Caption = strMainCaption
    f.lblSub.Caption = strSubCaption
    If lngSubMax = 0 And strSubCaption = "Sub Process" Then
        f.ProgressBarSub.Visible = False
        f.lblSub.Visible = False
    Else
        f.ProgressBarSub.Max = lngSubMax
        f.ProgressBarSub.Visible = True
        f.lblSub.Visible = True
    End If

ExitHere:
  Exit Function

HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.StartProgressBars"
       Resume ExitHere
  End Select
  Resume 'Debug only
End Function

Public Function UpdateProgressBars(Optional lngMain As Long = 0, _
    Optional lngSub As Long = 0, _
    Optional lngSubMax As Long = 0)
    
    Dim f As Access.Form


On Error GoTo HandleErr

    If Not IsLoaded("frmProgressBar") Then Exit Function
    Set f = Forms![frmProgressBar]
    If lngSubMax > 0 Then f.ProgressBarSub.Max = lngSubMax
    If lngMain > 0 Then f.ProgressBarMain.Value = lngMain
    If lngSub > 0 Then f.ProgressBarSub.Value = lngSub
    DoEvents

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.UpdateProgressBars"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Function CloseProgressBars()

On Error GoTo HandleErr

    If Not IsLoaded("frmProgressBar") Then Exit Function
    DoCmd.Close acForm, "frmProgressBar", acSaveYes

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.CloseProgressBars"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Function RemoveAmpersand(strCaption) As Ampersand
    Dim intPosn As Integer 'Added 17/06/2004
    intPosn = InStr(strCaption, "&")
    If intPosn = 0 Then
        RemoveAmpersand.strValue = strCaption
        RemoveAmpersand.intPosnOfAmpersand = 0
    Else
        RemoveAmpersand.strValue = Left(strCaption, intPosn - 1) & Mid(strCaption, intPosn + 1)
        RemoveAmpersand.intPosnOfAmpersand = intPosn
    End If
    
End Function

Public Sub DumpAllMenus(Optional TopMenu As String = "Menu Bar")

    ' Dump a list of all menus and their sub-items to
    ' the Debug window. This list can be quite lengthy.
    ' This procedure walks its way through the Menu Bar
    ' commandbar object, only.
    ' To look at just a single menu bar item (File,
    ' Edit, Window, etc.), pass that value as
    ' this procedure's only parameter.
    '
    ' NOTE: This code will only work with the English
    ' version of Access, because it uses the menu name
    ' to index into the top-level menu bar.
    
    ' In:
    '    TopMenu: (optional) A top-level menu bar name
    '     such as "File", "Edit", "Help", "Window"
    '     If you pass nothing, the procedure walks through
    '     all the menus.
    
    ' From Access 2002 Desktop Developer's Handbook
    ' by Litwin, Getz, Gunderloy (Sybex)
    ' Copyright 2001.  All rights reserved.
    
    Dim cbr As CommandBar
    Dim cbp As CommandBarPopup
    
    Set cbr = CommandBars("Menu Bar")
    If TopMenu <> "Menu Bar" Then
        Set cbp = cbr.Controls(TopMenu)
        Call DumpMenu(cbp, 1)
    Else
        For Each cbp In cbr.Controls
            'Debug.Print cbp.Caption
            Call DumpMenu(cbp, 1)
        Next cbp
    End If
End Sub

Private Sub DumpMenu(cbp As CommandBarPopup, intLevel As Integer)
    ' Called from DumpAllMenus, or recursively from DumpMenu,
    ' to dump information about a specific commandbar.
    Dim cbc As commandBarControl
    Dim intI As Integer

    ' From Access 2002 Desktop Developer's Handbook
    ' by Litwin, Getz, Gunderloy (Sybex)
    ' Copyright 2001.  All rights reserved.

    For Each cbc In cbp.CommandBar.Controls
        ' Insert enough spaces to indent according to the
        ' level of recursion.
        For intI = 0 To intLevel
            'Debug.Print "   ";
        Next intI
        'Debug.Print cbc.Caption, cbc.id
        If cbc.Type = msoControlPopup Then
            ' Call this routine recursively, to document
            ' the next lower level.
            Call DumpMenu(cbc.Control, intLevel + 1)
        End If
    Next cbc
End Sub



Private Function LaunchExplorer()
    Dim strAppName As String 'Added 17/03/2004
    Dim ctl As Control
    On Error GoTo HandleErr
    Set ctl = Screen.activeControl
    If ctl.ControlType = acTextBox Then
        strAppName = "EXPLORER.EXE /e,/select," & ctl.Value
        Call shell(strAppName, 1)
    End If


ExitHere:
  Exit Function

HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basImportFiles.LaunchExplorer"
       Resume ExitHere
  End Select
  Resume 'Debug only
End Function

Public Function DoTabOrder(strObject As String, acObject As Access.AcObjectType _
    ) As Boolean
    
    Dim f As Access.Form
    Dim sfm As Access.Form
    Dim ctl As Access.Control
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    Dim ctl2 As Access.Control
    Dim ctl3 As Access.Control
    Dim ctlsub As Access.Control
    Dim strTemp As String
    On Error GoTo HandleErr
    Set db = codeDB
    Set rs = db.OpenRecordset("tmpTabOrder")
    DoCmd.Close acForm, "frmTabOrder", acSaveYes
    ZapTable rs
    DoTabOrder = False
    If acObject <> acForm Then Exit Function
    DoCmd.OpenForm strObject, acDesign
    Set f = Forms(strObject)
    For Each ctl In f.Controls
        strTemp = ""
        Select Case ctl.Section
        Case acHeader
            strTemp = "Header"
        Case acDetail
            strTemp = "Detail"
        Case acFooter
            strTemp = "Footer"
        Case acPageHeader  'Tabs are of no use here
            strTemp = "PageHeader"
        Case acPageFooter 'Tabs are of no use here
            strTemp = "PageFooter"
        End Select
        If Not strTemp = "PageHeader" And Not strTemp = "PageFooter" Then
            Select Case ctl.ControlType
            Case acToggleButton, acCheckBox, acOptionButton, acCommandButton, acSubform, acBoundObjectFrame, acComboBox, acListBox _
                , acObjectFrame, acOptionGroup, acTabCtl, acTextBox, acPage
                
                If ctl.Parent.Name <> strObject Then 'Controls not applicable if in a container
                    Set ctl2 = f.Controls(ctl.Parent.Name)
                    If ctl2.ControlType = acPage Then
                        rs.AddNew
                        rs![Parent] = ctl.Parent.Name
                        rs![Section] = strTemp
                        rs![ControlName] = ctl.Name
                        rs![FormName] = strObject
                        rs![TabIndex] = ctl.TabIndex
                        rs![Type] = ctl.ControlType
                        rs![ParentType] = ctl2.ControlType
                        rs![Left] = ctl.Left
                        rs![Top] = ctl.Top
                        rs.Update
                    ElseIf ctl2.ControlType = acOptionGroup Then
                        'Tabs not applicable
                    ElseIf ctl2.ControlType = acTabCtl Then
                        'Tabs not applicable
                    End If
                Else
                    rs.AddNew
                    rs![Parent] = ctl.Parent.Name
                    rs![Section] = strTemp
                    rs![ControlName] = ctl.Name
                    rs![FormName] = strObject
                    rs![TabIndex] = ctl.TabIndex
                    rs![Type] = ctl.ControlType
                    rs![Left] = ctl.Left
                    rs![Top] = ctl.Top
                    rs.Update
                End If
                
            End Select
        End If
    Next ctl
    rs.Close
    DoCmd.OpenForm "frmTabOrder", acNormal, , _
    "[Section] = 'Detail' And [Parent] = '" & strObject & "'" _
    , acFormEdit, acWindowNormal, strObject
    DoTabOrder = True

ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case # '
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.CheckAccelerators"   'ErrorHandler:$$N=basCommon.CheckAccelerators
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function


' --------------------------------------------------------------------------
' Function: MkTree
' Purpose : Create directories for a given file name
' Notes   : MkTree "c:\a\b\c\file.txt" would create directories for c:\a\b\c
'           MkTree "c:\a\b\c\" would create c:\a\b\c
'           MkTree "c:\a\b\c" would create c:\a\b
' --------------------------------------------------------------------------
Public Sub MkTree(ByVal strFile As String)
    ' walk the file and make sure the paths exist
    Dim strRoot As String
    Dim strPath As String
    Dim pos As Integer

    ' get the root
    If (InStr(strFile, ":\") = 2) Then
        strRoot = Left(strFile, InStr(strFile, ":\") + 1)
    ElseIf (InStr(strFile, "\\") = 1) Then
        strRoot = Left(strFile, InStr(InStr(strFile, "\\") + 2, strFile, "\"))
    Else
        MsgBox "Invalid Root Directory", vbExclamation
        Exit Sub
    End If

    pos = InStr(Len(strRoot) + 1, strFile, "\")
    While (pos > 0)
        strPath = Left(strFile, pos)

        ' Create the directory
        On Error Resume Next
        MkDir strPath
        Debug.Assert Err = 0 Or Err = 75
        On Error GoTo 0
        pos = InStr(pos + 1, strFile, "\")
    Wend
End Sub

Public Function ISColumnEmpty(strTable As String, strColumn As String) As Boolean
    Dim strSQLtext As String
    Dim db As DAO.database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    strSQLtext = "SELECT [" & strTable & "].[" & strColumn & "]" & vbCrLf
    strSQLtext = strSQLtext & "        FROM [" & strTable & "]" & vbCrLf
    strSQLtext = strSQLtext & "       WHERE (((Len([" & strTable & "]![" & strColumn & "] & """"))>0))" & vbCrLf
    strSQLtext = strSQLtext & "    GROUP BY [" & strTable & "].[" & strColumn & "];"
    
    Set rs = db.OpenRecordset(strSQLtext, dbOpenSnapshot)
    If Not rs.EOF Then
        ISColumnEmpty = False
    Else
        ISColumnEmpty = True
    End If

    rs.Close

End Function


'Public Function DoesFileExists(strFile As String) As Boolean
'    ' Return existance of file based on
'    ' a directory search.
'
'On Error GoTo HandleErr
'
'    On Error Resume Next
'    DoesFileExists = (Len(Dir$(strFile)) > 0)
'    On Error GoTo HandleErr
'    If Err.Number <> 0 Then
'        DoesFileExists = False
'    End If
'
'ExitHere:
'  Exit Function
'
'
'' Automatic error handler last updated at 10 March 2004 11:33:31
'HandleErr:
'  Select Case Err.Number
'  'Case # '
'  Case Else
'       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basImportFiles.DoesFileExists"
'       Resume ExitHere
'  End Select
'  Resume 'Debug only
'
'End Function


Public Function AddTranslations(strObject As String, acObject As Access.AcObjectType) As Boolean
    Dim currentDatabase As DAO.database
    Dim rs As DAO.Recordset
    Dim recordsetTranslations As DAO.Recordset
    Dim accessControl As Access.Control
    Dim aobCurrent As Object
    Dim strSQLtext As String
    Dim strMsg As String, strHeading As String
    Dim intCounter As Integer, strControl As String
    Dim strNewName As String
    Dim strTemp As String, k As Integer
    AddTranslations = False
    On Error GoTo HandleErr
    Select Case acObject
    Case acForm
        DoCmd.OpenForm strObject, acDesign
        Set aobCurrent = Forms(strObject)
    Case acReport
        DoCmd.OpenReport strObject, acDesign
        Set aobCurrent = Reports(strObject)
    End Select
    
    Set currentDatabase = CurrentDb()
    Dim Recordset As DAO.Recordset
    Set Recordset = currentDatabase.OpenRecordset("Translations Lookup")
    Set recordsetTranslations = currentDatabase.OpenRecordset("Translations")
    For Each accessControl In aobCurrent.Controls
        Dim ControlType As Access.AcControlType
        ControlType = accessControl.ControlType
        If ControlType = acCommandButton Or ControlType = acLabel Or ControlType = acNavigationButton Or ControlType = acToggleButton Then
                intCounter = intCounter + 1
                Recordset.AddNew
                Recordset![English] = accessControl.Caption
                Recordset.Update
                recordsetTranslations.AddNew
                recordsetTranslations![Control] = Left(accessControl.Name, 255)
                recordsetTranslations![object] = strObject
                recordsetTranslations![Property Name] = "Caption"
                recordsetTranslations![English] = accessControl.Caption
                recordsetTranslations![Object Type] = IIf(acObject = acForm, "Form", "Report")
                recordsetTranslations.Update
        End If
    Next
    Recordset.Close: Set Recordset = Nothing: Set currentDatabase = Nothing
    DoCmd.OpenTable "Translations Lookup", acViewNormal, acEdit
    AddTranslations = True
ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    Case 3022 ' update cancelled due to unique constraint
         Resume Next
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.AddTranslations"  'ErrorHandler:$$N=basCommon.AddTranslations
        Resume ExitHere
    End Select
Resume 'Debug Only
End Function


Public Sub PutCodeandProceduresinTable()
    Dim k As Integer
    Dim module As module
    Dim lngCount As Long, lngCountDecl As Long, lngI As Long
    Dim strProcName As String, astrProcNames() As String
    Dim intI As Integer, strMsg As String
    Dim lngR As Long
    On Error GoTo HandleErr
        Dim currentDatabase As DAO.database
        Dim Recordset As DAO.Recordset
        Set currentDatabase = codeDB
        Set Recordset = currentDatabase.OpenRecordset("SearchCodeTable")
    Dim stringSQLText As String
    stringSQLText = "DELETE SearchCodeTable.ID" & vbCrLf
    stringSQLText = stringSQLText & "        FROM SearchCodeTable;"
        currentDatabase.Execute stringSQLText, dbFailOnError
        
    ' Open specified Module object.
    'DoCmd.OpenModule strModuleName
    ' Return reference to Module object.
'    Debug.Print Application.Modules.Count
    SysCmd acSysCmdInitMeter, "Please wait while this process runs…", Modules.Count
    Dim AccessObject As AccessObject
    For Each AccessObject In CurrentProject.AllModules
    'For k = 0 To Modules.Count - 1
        k = k + 1
        SysCmd acSysCmdUpdateMeter, k
       ' Set module = Modules(k)
       DoCmd.OpenModule AccessObject.FullName
        Set module = Application.Modules(AccessObject.FullName)
        ' Count lines in module.
        lngCount = module.CountOfLines
        ' Count lines in Declaration section in module.
        lngCountDecl = module.CountOfDeclarationLines
        ' Determine name of first procedure.
        strProcName = module.ProcOfLine(lngCountDecl + 1, lngR)
        ' Initialize counter variable.
        intI = 0
        ' Redimension array.
'        ReDim Preserve astrProcNames(intI)
        ' Store name of first procedure in array.
'        astrProcNames(intI) = strProcName
        ' Determine procedure name for each line after declarations.
        For lngI = 1 To lngCount
            ' Compare procedure name with ProcOfLine property value.
                strProcName = module.ProcOfLine(lngI, lngR)
            
            Recordset.AddNew
            Recordset![ModuleName] = AccessObject.FullName
            Recordset![ProcedureName] = strProcName
            Recordset![Code] = module.Lines(lngI, 1)
            Recordset![line] = lngI
            Recordset.Update
           DoCmd.Close acModule, AccessObject.FullName, acSaveYes
        Next lngI
                 
    Next
    Recordset.Close
    SysCmd acSysCmdRemoveMeter
    
ExitHere:
    Exit Sub

HandleErr:
    Select Case Err.Number
    Case 2017 'VBA Code protected - please supply a password
        MsgBox "The VBA Code is protected - please unlock by supplying a password first", vbInformation, "Code Protected"
        DoCmd.RunCommand acCmdVisualBasicEditor
        Resume ExitHere
    Case 2517  'Access can't find the procedure
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
        Resume ExitHere
    End Select
    Resume 'Debug Only
End Sub