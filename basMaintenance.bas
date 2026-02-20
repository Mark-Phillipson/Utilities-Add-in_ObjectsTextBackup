Option Compare Database
Option Explicit
Const adhcErrActionCancelled = 2501

' Set to False to include the current form in the list of forms,
' or set to True to not show the current form.
Const adhcSkipThisForm = True

Private Function IsHidden( _
 aob As AccessObject) As Boolean

    ' Determine whether or not the specified object is
    ' hidden in the Access database window


    On Error GoTo HandleErr

     If Application.GetHiddenAttribute( _
     aob.Type, aob.Name) Then
        IsHidden = True
    End If

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.IsHidden"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function

Private Function IsDeleted( _
 ByVal strName As String) As Boolean

    On Error GoTo HandleErr

    IsDeleted = (Left(strName, 7) = "~TMPCLP")

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.IsDeleted"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function

Private Function IsSystemObject( _
 aob As AccessObject) As Boolean

    ' Determine whether or not the specified object is
    ' an Access system object or not.


    On Error GoTo HandleErr

    Const conSystemObject = &H80000000
    Const conSystemObject2 = &H2
    
    If (Left$(aob.Name, 4) = "USys") Or _
     Left$(aob.Name, 4) = "~sq_" Then
        IsSystemObject = True
    Else
        If (aob.Attributes And conSystemObject) = _
         conSystemObject Then
            IsSystemObject = True
        Else
            If (aob.Attributes And conSystemObject2) = _
             conSystemObject2 Then
                IsSystemObject = True
            End If
        End If
    End If

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.IsSystemObject"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function


Private Function GetObjectList( _
 ByVal lngType As AcObjectType) As String

    ' Returns a string with a semi-colon delimited list of object names.
    
    ' Parameters:
    '   intType -- one of acTable, acQuery, acForm,
    '              acReport, acDataAccessPage, acMacro or acModule
    
    Dim intI As Integer
    Dim fSystemObj As Boolean
    Dim strName As String
    Dim fShowHidden As Boolean
    Dim fIsHidden As Boolean
    Dim strOutput As String
    Dim fShowSystem As Boolean
    Dim objCollection As Object
    Dim aob As AccessObject

    On Error GoTo HandleErrors
    DoCmd.Hourglass True
    
    ' Are you supposed to show hidden/system objects?
    fShowHidden = _
     Application.GetOption("Show Hidden Objects")
    fShowSystem = _
     Application.GetOption("Show System Objects")
    
    Select Case lngType
        Case acTable
            Set objCollection = CodeData.AllTables
        Case acQuery
            Set objCollection = CodeData.AllQueries
        Case acForm
            Set objCollection = CodeProject.AllForms
        Case acReport
            Set objCollection = CodeProject.AllReports
        Case acDataAccessPage
            Set objCollection = _
             CodeProject.AllDataAccessPages
        Case acMacro
            Set objCollection = CodeProject.AllMacros
        Case acModule
            Set objCollection = CodeProject.AllModules
    End Select
            
    For Each aob In objCollection
        fIsHidden = IsHidden(aob)
        strName = aob.Name
        fSystemObj = IsSystemObject(aob)
        ' Unless this is a system object and
        ' you're not showing system objects...
        If (fSystemObj Imp fShowSystem) Then
            ' If the object isn't deleted and its hidden
            ' characteristics match those you're
            ' looking for...
            If Not IsDeleted(strName) And _
             (fIsHidden Imp fShowHidden) Then
                ' If this isn't a form, just add it to
                ' the list. If it is, one more check:
                ' is this the CURRENT form? If so, and if
                ' the flag isn't set to include the current
                ' form, then skip it.
                Select Case lngType
                    Case acForm
                        If Not (adhcSkipThisForm And _
                         (strName = "frmMaintenance")) Then
                            strOutput = _
                            strOutput & ";" & strName
                        End If
                    Case Else
                        strOutput = _
                         strOutput & ";" & strName
                End Select
            End If
        End If
    Next aob
    strOutput = Mid$(strOutput, 2)
    
ExitHere:
    DoCmd.Hourglass False
    GetObjectList = strOutput
    Exit Function

HandleErrors:
    HandleErrors Err.Number, "GetObjectList"
    Resume ExitHere
End Function
Private Sub HandleErrors(intErr As Integer, strRoutine As String)

    On Error GoTo HandleErr

    MsgBox "Error: " & Error(intErr) & " (" & intErr & ")", vbExclamation, strRoutine

ExitHere:
  Exit Sub

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.HandleErrors"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Sub


Public Function dbmGetTblNames(strDBPath As String, blnShowAllTypes As Boolean _
    , blnShowHidden As Boolean, Optional blnShowLinked As Boolean) As Boolean
    
    ' Fill tmpTableNames with the table names of the strDBPath
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
    Dim lngKounter As Long
    On Error GoTo dbmGetTblNames_Err
    dbmGetTblNames = True
    
    fShowSystem = False
    fShowHidden = blnShowHidden
    adhcTrackingHidden = True
    
    Set db = codeDB
    Set dbData = DBEngine.Workspaces(0).OpenDatabase(strDBPath)
    
    Set rs = db.OpenRecordset("tmpTableNames", dbOpenTable)
    
    v = ZapTable(rs)
            SysCmd acSysCmdInitMeter, "Filling list with tables", dbData.TableDefs.Count
            dbData.TableDefs.Refresh
            For Each tdf In dbData.TableDefs
                ' Check and see if this is a system object.
                fSystemObj = dbmisSystemObject(acTable, tdf.Name, tdf.Attributes)
                If adhcTrackingHidden Then
                    If strDBPath = CurrentDb.Name Then
                        fIsHidden = Application.GetHiddenAttribute(acTable, tdf.Name)
                    Else
                        fIsHidden = False
                    End If
                Else
                    ' If not tracking hidden objects,
                    ' just assume it's not hidden!
                    fIsHidden = False
                End If
                ' Unless this is a system object and you're not showing system
                ' objects, or this table has its hidden bit set,
                ' add it to the list.
                If (fSystemObj Imp fShowSystem) And _
                 ((tdf.Attributes And dbHiddenObject) = 0) _
                 And (fIsHidden Imp fShowHidden) Then
                    ' do not include linked tables
                    If tdf.Attributes <> dbAttachedTable Or blnShowLinked Then
                    
                        blnShow = True
                        If blnShowAllTypes Then
                            blnShow = True
                        Else
                            Select Case Left(tdf.Name, 3)
'                            Case "tmp", "tbl", "Zsy"
'                                blnShow = False

                            Case Else
                                If tdf.Name = "prnDesignNames" Then
                                    blnShow = False
                                ElseIf tdf.Name Like "Test*" Then
                                    blnShow = False
                                ElseIf tdf.Name = "Blank Carriers" Then
                                    blnShow = False
                                ElseIf tdf.Name = "Switchboard Items" Then
                                    blnShow = False
                                End If
                            End Select
                        End If
                        
                        
                        If blnShow Then
                            rs.AddNew
                            rs![TableName] = tdf.Name
                            rs![Last Updated] = tdf.LastUpdated
                            rs.Update
                        End If
                    End If

                End If
                lngKounter = lngKounter + 1
                SysCmd acSysCmdUpdateMeter, lngKounter
            Next tdf
    If Not rs.EOF Then rs.MoveLast
    
    Forms![frmMaintenance]![txtQty] = rs.RecordCount
    Forms![frmQtyAnalizer]![txtQty] = rs.RecordCount
    SysCmd acSysCmdRemoveMeter
dbmGetTblNames_Exit:
    rs.Close
    dbData.Close
    Exit Function
dbmGetTblNames_Err:
    
    If Err.Number = 3044 Then
        MsgBox "Please enter a valid database path and name", vbExclamation
        dbmGetTblNames = False
    ElseIf Err.Number = 3024 Then 'Could not find file
        MsgBox "Please check the file exists;" & vbCrLf & Err.Description, vbExclamation, "File not Found"
        Resume dbmGetTblNames_Exit
    ElseIf Err.Number = 2450 Then  'Table Maintenance can't find the form 'frmMaintenance' referred to in a macro expression or Visual Basic code.
        Resume Next
    Else
        MsgBox Err.Number & vbCrLf & vbCrLf & Err.Description
        dbmGetTblNames = False
        GoTo dbmGetTblNames_Exit
    End If
    
Resume
End Function

Function dbmisSystemObject(intType As Integer, ByVal strName As String, _
 Optional ByVal varAttribs As Variant)

    ' Determine whether or not the object named 'strName' is
    ' an Access system object or not.


    On Error GoTo HandleErr

    If IsMissing(varAttribs) Then
        varAttribs = 0
    End If
    
    If (Left$(strName, 4) = "USys") Or Left$(strName, 4) = "~sq_" Then
        dbmisSystemObject = True
    Else
        dbmisSystemObject = ((intType = acTable) And _
         ((varAttribs And dbSystemObject) <> 0))
    End If

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.dbmisSystemObject"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function

Public Function TblMaintEntry()

    On Error GoTo HandleErr

    DoCmd.OpenForm "frmMain"

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.TblMaintEntry"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function

Public Function ZapTable(rs As DAO.Recordset) As Boolean

    On Error GoTo HandleErr

    If Not rs.EOF Then
        rs.MoveFirst
    End If
    Do Until rs.EOF
        rs.Delete
        rs.MoveNext
    Loop

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.ZapTable"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function

Public Function ToggleHiddenProperty(Optional strTable As String, Optional blnShow As Boolean)
    
On Error GoTo HandleErr
    Dim blnHidden As Boolean
    
    If Len(strTable) = 0 Then
        If Application.CurrentObjectType <> acTable And Application.CurrentObjectType <> acDefault Then Exit Function
        strTable = Application.CurrentObjectName
    End If

    If blnShow Then
        Application.SetHiddenAttribute acTable, strTable, False
    Else
        blnHidden = Application.GetHiddenAttribute(acTable, strTable)
        Application.SetHiddenAttribute acTable, strTable, Not blnHidden
    End If
    Exit Function

    Application.RefreshDatabaseWindow
ExitHere:
    Exit Function

' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 12-28-2000 08:59:30   'ErrorHandler:$$D=12-28-2000    'ErrorHandler:$$T=08:59:30
HandleErr:
    Select Case Err.Number
    Case 3265 'Item cannot be found in the collection corresponding to the requested name or ordinal.
        MsgBox "Cannot find table: " & strTable, vbExclamation, "Hide/Show Table..."
        Resume ExitHere
    Case 3011 ' could not find the object make sure it exists
    MsgBox "The object could not be found in the database." _
    , vbExclamation + vbOKOnly + vbDefaultButton1 _
    , "Could Not Find Object"
        Resume ExitHere
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "basTableProcessing.ToggleHiddenProperty"    'ErrorHandler:$$N=basTableProcessing.ToggleHiddenProperty
        Resume ExitHere
    End Select
' End Error handling block.
End Function


Public Function ImportSelectedTable()
        Dim frm As Form
        Dim ctl As Control
        Dim strTableName As String
        Dim varItm As Variant
        

    On Error GoTo HandleErr

        Set frm = Forms![frmMaintenance]
        If frm.cboDBName = CurrentDb.Name Then
            Exit Function
        End If
        Set ctl = frm![lstTables]
        If ctl.ItemsSelected.Count = 0 Then
            Exit Function
        End If
        For Each varItm In ctl.ItemsSelected
            strTableName = ctl.Column(0, varItm)
            If Len(Nz(strTableName, "")) > 0 Then
                If DoesTableExits(strTableName) Then
                    Select Case FormattedMsgBox(strTableName & "@The table already exists in this database.@Do you want to replace the existing table?", vbYesNo + vbExclamation + vbDefaultButton2, "Import Table")
                    Case vbYes
                        DoCmd.DeleteObject acTable, strTableName
                    Case vbNo
                        Exit Function
                    End Select
                End If
                'Import table
                DoCmd.TransferDatabase acImport, "Microsoft Access", frm.cboDBName, acTable, strTableName, strTableName, False
            End If
        Next varItm
        DoCmd.SelectObject acTable, strTableName, True


ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.ImportSelectedTable"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function

Public Function DoesTableExits(strTable As String) As Boolean
    Dim db As DAO.database
    Dim intCounter As Integer

    On Error GoTo HandleErr

    Set db = CurrentDb
    DoesTableExits = False
    For intCounter = 0 To db.TableDefs.Count - 1
        If strTable = db.TableDefs(intCounter).Name Then
            DoesTableExits = True
            Exit For
        End If
    Next
    

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 04 August 2006 16:11:54
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "An Error has occured please inform IT developer support " & Err.Number & ": " & Err.Description, vbCritical, "basMaintenance.DoesTableExits"
        Resume ExitHere
  End Select
'Debug Only
Resume
' End Error handling block.
End Function