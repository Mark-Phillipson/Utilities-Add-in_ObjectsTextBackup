   Option Compare Database
Option Explicit
#If Win64 Then
    Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Public Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

Enum MapVendor
    Google
    Bing
End Enum
    
Public Enum BasicDataType
    bdtText
    bdtnumber
    bdtdate
End Enum

Public Enum MeasurementConversion
    totwips
    ToCentimetres
End Enum



Function IsLoaded(MyFormName)
    '  Determines if a form is loaded.

On Error GoTo HandleErr

    Const FORM_DESIGN = 0
    Dim i As Integer
    IsLoaded = False
    For i = 0 To Forms.Count - 1
        If Forms(i).FormName = MyFormName Then
            If Forms(i).CurrentView <> FORM_DESIGN Then
                IsLoaded = True
                Exit Function  '  Quit function once form has been found.
            End If
        End If
    Next

ExitHere:
  Exit Function

' Error handling block added by VBA Code Commenter and Error Handler Add-In. DO NOT EDIT this block of code.
' Automatic error handler last updated at 11 May 2004 10:59:38
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basCommon.IsLoaded"
       Resume ExitHere
  End Select
  Resume 'Debug only
' End Error handling block.
End Function

Public Function GetDirectory()
 Dim strDir As String
On Error GoTo HandleErr
        Dim strFileName  As String
        With Application.FileDialog(msoFileDialogFolderPicker)
            '.Filters.Clear
            '.'Filters.Add "Picture File", "*.bmp"
            .InitialView = msoFileDialogViewList
            .Title = "Choose a Directory"
            If .Show Then
                strDir = .SelectedItems(1)
            End If
        End With
   ' Choose a Directory    strDir = GetFileNameOfficeDialog(True, True, False, True, True, 0, "", 1, "", "Choose a Directory", "Select", "C:\", True)
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
 Dim stringFilename As String
On Error GoTo HandleErr
        Dim strFileName  As String
        With Application.FileDialog(msoFileDialogOpen)
            .Filters.Clear
            .Filters.Add "Microsoft Word", "*.docx"
            .Filters.Add "All Files", "*.*"
            .InitialView = msoFileDialogViewList
            .Title = "Choose a File"
            If .Show Then
                stringFilename = .SelectedItems(1)
            End If
        End With
   ' Choose a Directory    strDir = GetFileNameOfficeDialog(True, True, False, True, True, 0, "", 1, "", "Choose a Directory", "Select", "C:\", True)
    If Len(stringFilename) > 0 Then
        Screen.activeControl = stringFilename
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

Public Function ConvertDatetoUSFormat(stringUKDate As String) As String
     If Len(stringUKDate & "") = 10 Then
        ConvertDatetoUSFormat = Mid(stringUKDate, 4, 3) & Left(stringUKDate, 3) & Right(stringUKDate, 4)
    End If
End Function

'Public Function ConvertDelimitedToArray(strDelimiter As String _
'    , strValue As String) As Variant
'    Dim intDelimiters As Integer 'Added 14/10/2005
'    Dim strTemp As String
'    Dim intPosn As Integer
'    Dim k As Integer
'    Dim varArray() As Variant
'    'ReDim Preserve X(10, 10, 15)
'    'Count number of delimiters
'    intDelimiters = 0
'    For k = 1 To Len(strValue)
'        If Mid(strValue, k, Len(strDelimiter)) = strDelimiter Then
'            intDelimiters = intDelimiters + 1
'        End If
'    Next
'    intDelimiters = intDelimiters + 1
'    ReDim varArray(intDelimiters) As Variant
'    k = -1
'    If Len(strValue & "") > 0 And InStr(strValue, strDelimiter) > 0 Then
'        strTemp = strValue
'        Do Until InStr(strTemp, strDelimiter) = 0
'            intPosn = InStr(strTemp, strDelimiter)
'            k = k + 1
'
'            varArray(k) = Left(strTemp, intPosn - 1)
'            strTemp = Mid(strTemp, intPosn + Len(strDelimiter))
'        Loop
'        varArray(k + 1) = strTemp
'
'    End If
'    ConvertDelimitedToArray = varArray
'End Function

Public Sub Zap(stringTableName As String)
    Dim stringSQLText As String
    stringSQLText = "DELETE " & vbCrLf
    stringSQLText = stringSQLText & "        FROM [" & stringTableName & "];"
    CurrentDb.Execute stringSQLText, dbFailOnError
End Sub

Public Function DoesFileExists(strFile As String) As Boolean
    ' Return existance of file based on
    ' a directory search.

On Error GoTo HandleErr

    On Error Resume Next
    DoesFileExists = (Len(Dir$(strFile)) > 0)
    On Error GoTo HandleErr
    If Err.Number <> 0 Then
        DoesFileExists = False
    End If

ExitHere:
  Exit Function


' Automatic error handler last updated at 10 March 2004 11:33:31
HandleErr:
  Select Case Err.Number
  'Case # '
  Case Else
       MsgBox "Unexpected Error Has Occured Please Inform IT Support " & Err.Number & ": " & Err.Description, vbCritical, "basImportFiles.DoesFileExists"
       Resume ExitHere
  End Select
  Resume 'Debug only

End Function

'Public Function atCNames(UOrC As Integer) As String
'
'If UOrC = 1 Then
'    atCNames = Environ("USERNAME")
'Else
'    atCNames = Environ("COMPUTERNAME")
'End If
'
'End Function

Public Function GetWindowsUsername() As String
    GetWindowsUsername = Environ("USERNAME")
End Function

Public Function DoesFolderExist(stringFolder As String) As Boolean
    Dim stringTemporary As String
    Dim blnResult As Boolean
    Dim FilesSystemObject    As Object
    Set FilesSystemObject = CreateObject("Scripting.FileSystemObject")
    blnResult = FilesSystemObject.FolderExists(stringFolder)
    DoesFolderExist = blnResult
    

ExitHere:
   Exit Function

DoesFolderExist_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure DoesFolderExist of Module General" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function


Public Function RemoveIllegalFileNameCharacters(stringValue As String, Optional stringReplaceWith As String = "") As String
    '< > : " / \ | ? *
    Dim stringNew As String
    If Len(stringValue) = 0 Then
        Exit Function
    End If
    Dim integerCounter As Integer
    For integerCounter = 1 To Len(stringValue)
        Select Case Mid(stringValue, integerCounter, 1)
        Case "<", ">", ":", """", "/", "\", "|", "?", "*"
            stringNew = stringNew & stringReplaceWith
        Case Else
            stringNew = stringNew & Mid(stringValue, integerCounter, 1)
        End Select
    Next
    RemoveIllegalFileNameCharacters = stringNew
        
End Function


Public Sub MapToLocation(stringCompleteAddress As String, vendor As MapVendor)
    Dim stringHyperlink As String
    On Error GoTo MapToLocation_Error

    If vendor = Bing Then
        stringHyperlink = "http://www.bing.com/maps/default.aspx?setmkt=en-US&where1=" & stringCompleteAddress
    ElseIf vendor = Google Then
        stringHyperlink = "https://maps.google.com/?q=" & stringCompleteAddress
    End If
        
    Application.FollowHyperlink stringHyperlink, , True, True

ExitHere:
   Exit Sub

MapToLocation_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure MapToLocation of Module General" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only

    
End Sub
Public Function ManageReference(BooleanAddReference As Boolean)
    Dim stringAddin  As String
    
    On Error GoTo ManageReference_Error
    stringAddin = "C:\MSPSystems\Utilities Add-in.accdb"
    If BooleanAddReference Then
        If DoesFileExists(stringAddin) Then
            Application.References.AddFromFile stringAddin
            MsgBox "The reference has successfully been added." _
            , vbInformation + vbOKOnly + vbDefaultButton1 _
            , "Add-In Reference"
            ChangeCode BooleanAddReference
        End If
    Else
        Application.References.Remove Application.References("Utilities Add-in")
        MsgBox "The reference has successfully been removed." _
        , vbInformation + vbOKOnly + vbDefaultButton1 _
        , "Add-In Reference"
        ChangeCode BooleanAddReference
    End If
    ' Make a small change to the code module so the reference setting will persist
    On Error Resume Next
    DoCmd.Save

ExitHere:
   Exit Function

ManageReference_Error:
    Select Case Err.Number
        Case 9 ' subscript out of range  (reference has already been removed)
        Resume ExitHere
    Case 32813 ' name conflicts with existing module, project or object library ( the reference already exists)
        Resume ExitHere

    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure ManageReference of Module General" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function

Public Sub ChangeCode(BooleanAddReference As Boolean)
    On Error Resume Next
    DoCmd.OpenModule "General", "ManageReference"
    Application.VBE.ActiveCodePane.CodeModule.InsertLines 2, "'" & IIf(BooleanAddReference, "Reference Added ", "Reference Removed ") & Format(Now(), "YYYY/mm/dd hh:nn")
    On Error GoTo 0
End Sub


Public Function atCNames(UOrC As Integer) As String
'**************************************************
'Purpose:  Returns the User LogOn Name or ComputerName
'Accepts:  UorC; 1=User, anything else = computer
'Returns:  The Windows Networking name of the user or computer
'**************************************************
On Error Resume Next
Dim NBuffer As String
Dim Buffsize As Long
Dim Wok As Long
    
Buffsize = 256
NBuffer = Space$(Buffsize)
#If Win64 Then
    If UOrC = 1 Then
        Wok = GetUserName(NBuffer, Buffsize)
        atCNames = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
        Wok = GetComputerName(NBuffer, Buffsize)
        atCNames = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    End If
#Else
    If UOrC = 1 Then
        Wok = api_GetUserName(NBuffer, Buffsize)
        atCNames = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
        Wok = api_GetComputerName(NBuffer, Buffsize)
        atCNames = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    End If
#End If

End Function


Public Function LaunchExplorer(stringFile As String)
    Dim strAppName As String 'Added 17/03/2004
    On Error GoTo HandleErr
        strAppName = "EXPLORER.EXE /e,/select," & stringFile
        Call shell(strAppName, 1)


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


' Function to check that all field values exist between two tables
'example call for this function:
'stringSQLText = "SELECT DISTINCT Users.FullName" & vbCrLf
' stringSQLText = stringSQLText & "           , TemporaryLabourCharges.Owner" & vbCrLf
' stringSQLText = stringSQLText & "        FROM Users " & vbCrLf
' stringSQLText = stringSQLText & "  RIGHT JOIN TemporaryLabourCharges " & vbCrLf
' stringSQLText = stringSQLText & "          ON Users.FullName = TemporaryLabourCharges.Owner" & vbCrLf
' stringSQLText = stringSQLText & "       WHERE (((Users.FullName) Is Null));"
' If Not CheckAllExist(stringSQLText, "Owner", stringMessage) Then
'     MsgBox stringMessage _
'     , vbExclamation + vbOKOnly + vbDefaultButton1 _
'     , "User Name Is Missing"
' End If
Public Function CheckAllExist(stringSQLText As String, stringFieldName As String _
    , stringMessage As String) As Boolean
        Dim currentDatabase As DAO.database
        Dim Recordset As DAO.Recordset
        CheckAllExist = True
        Set currentDatabase = CurrentDb
        Set Recordset = currentDatabase.OpenRecordset(stringSQLText, dbOpenSnapshot)
        If Not Recordset.EOF Then
            stringMessage = "The following value/values exist in the import but do not exist in the database table:" & vbCrLf
            Recordset.MoveFirst
            CheckAllExist = False
        End If
        Dim stringExtra As String
        Do Until Recordset.EOF
            stringExtra = stringExtra & Recordset.Fields(stringFieldName).Value & vbCrLf
            Recordset.MoveNext
        Loop
        Recordset.Close
        If Not CheckAllExist Then
            stringMessage = stringMessage & stringExtra
        End If
        
End Function


Public Function StripCharacterFromString(stringValue As String, stringCharacterToRemove As String) As String
    Dim stringCurrentCharacter As String
    Dim longCounter As Long
    Dim stringTemporary As String
    For longCounter = 1 To Len(stringValue)
        stringCurrentCharacter = Mid(stringValue, longCounter, 1)
        If stringCurrentCharacter = stringCharacterToRemove Then
            ' do nothing
        Else
            stringTemporary = stringTemporary & stringCurrentCharacter
        End If
    Next
    StripCharacterFromString = stringTemporary
End Function

Public Function HexColour(strHex As String) As Long

    'converts Hex string to long number, for Colours
    'the leading # is optional

    'example usage
    'Me.iSupplier.BackColour = HexColour("FCA951")
    'Me.iSupplier.BackColour = HexColour("#FCA951")

    'the reason for this function is to programmatically use the
    'Hex Colours generated by the Colour picker.
    'The trick is, you need to reverse the first and last hex of the
    'R G B combination and convert to Long
    'so that if the Colour picker gives you this Colour #FCA951
    'to set this in code, we need to return CLng(&H51A9FC)
    
    Dim strR As String
    Dim strG As String
    Dim strB As String
    
    'strip the leading # if it exists
    If Left(strHex, 1) = "#" Then
        strHex = Right(strHex, Len(strHex) - 1)
    End If
    
    'reverse the first two and last two hex numbers of the R G B values
    strB = Right(strHex, 2)
    strG = Mid(strHex, 3, 2)
    strR = Left(strHex, 2)
    Dim red As Integer
    Dim green As Integer
    Dim blue As Integer
    red = CInt("&H" & strR)
    green = CInt("&H" & strG)
    blue = CInt("&H" & strB)
    
    
    HexColour = RGB(red, green, blue)

End Function

Public Function ConvertLinefeedToCarriageReturnandLinefeed(variantvalue As Variant) As Variant
' when importing in Excel data for example there maybe fields that contain just the linefeed character ASCII code 10.
' This does not display properly in Microsoft Access so this function will convert the single control character to  a carriage return and linefeed
    Dim stringValue As String
    If IsNull(variantvalue) Then
        Exit Function
    End If
    stringValue = Nz(variantvalue, "")
    Dim blnPreviousCarriageReturn As Boolean
    Dim integerCounter As Integer
    Dim stringCurrentCharacter As String
    Dim stringTemporary As String
    For integerCounter = 1 To Len(stringValue)
        stringCurrentCharacter = Mid(stringValue, integerCounter, 1)
        Select Case stringCurrentCharacter
        Case Chr(13) ' carriage return
            blnPreviousCarriageReturn = True
            stringTemporary = stringTemporary & stringCurrentCharacter
        Case Chr(10) ' linefeed
            If Not blnPreviousCarriageReturn Then
                stringTemporary = stringTemporary & vbCrLf
            Else
                stringTemporary = stringTemporary & stringCurrentCharacter
            End If
        Case Else
            stringTemporary = stringTemporary & stringCurrentCharacter
        End Select
    Next
    ConvertLinefeedToCarriageReturnandLinefeed = stringTemporary
End Function

Public Function FixQuotes(varValue As Variant)
    ' Double any quotes inside varValue, and
    ' surround it with quotes.
    FixQuotes = "'" & Replace$(varValue & "", "'", "''", Compare:=vbTextCompare) & "'"
End Function

Public Function FixQuotesInside(varValue As Variant)
    ' Double any quotes inside varValue, and
    ' surround it with quotes.
    FixQuotesInside = Replace$(varValue & "", "'", "''", Compare:=vbTextCompare)
End Function

Public Sub FillListwithObjectNames(Control As Control, ObjectType As Integer _
    , Optional stringTableName As String = "", Optional stringQueryName As String = "", Optional blnFields As Boolean)
    Dim currentDatabase As DAO.database
    Dim Recordset As DAO.Recordset
    Dim stringSQLText As String
    On Error GoTo FillListwithObjectNames_Error
    Set currentDatabase = CurrentDb
    Dim codeDatabase As DAO.database: Set codeDatabase = codeDB
    Dim DAOField As DAO.field
    stringSQLText = "temporaryAccessObjects"
    Set Recordset = codeDatabase.OpenRecordset(stringSQLText)
    Dim AccessObject As Object
    If ObjectType = acReport Then
        stringSQLText = "DELETE temporaryAccessObjects.ObjectType" & vbCrLf
        stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
        stringSQLText = stringSQLText & "       WHERE (((temporaryAccessObjects.ObjectType)=" & ObjectType & "));"
        codeDatabase.Execute stringSQLText, dbFailOnError
        For Each AccessObject In Application.CurrentProject.AllReports
            Recordset.AddNew
            Recordset![ObjectName] = AccessObject.Name
            Recordset![ObjectType] = ObjectType
            Recordset.Update
        Next
    ElseIf ObjectType = acForm Then
        stringSQLText = "DELETE temporaryAccessObjects.ObjectType" & vbCrLf
        stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
        stringSQLText = stringSQLText & "       WHERE (((temporaryAccessObjects.ObjectType)=" & ObjectType & "));"
        codeDatabase.Execute stringSQLText, dbFailOnError
        For Each AccessObject In Application.CurrentProject.AllForms
            Recordset.AddNew
            Recordset![ObjectName] = AccessObject.Name
            Recordset![ObjectType] = ObjectType
            Recordset.Update
        Next
    ElseIf ObjectType = acTable Then
        If stringTableName <> "" And blnFields Then
            stringSQLText = "DELETE temporaryAccessObjects.ObjectType" & vbCrLf
            stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
            stringSQLText = stringSQLText & "       WHERE (((temporaryAccessObjects.ObjectType)=" & 13 & "));"
            codeDatabase.Execute stringSQLText, dbFailOnError
            Set currentDatabase = CurrentDb
            Dim tableDefinition As DAO.tableDef
            Set tableDefinition = currentDatabase.TableDefs(stringTableName)
            For Each DAOField In tableDefinition.Fields
                Recordset.AddNew
                Recordset![ObjectName] = DAOField.Name
                Recordset![ObjectType] = 13: ObjectType = 13
                Recordset.Update
            Next
        Else
            stringSQLText = "DELETE temporaryAccessObjects.ObjectType" & vbCrLf
            stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
            stringSQLText = stringSQLText & "       WHERE (((temporaryAccessObjects.ObjectType)=" & ObjectType & "));"
            codeDatabase.Execute stringSQLText, dbFailOnError
            For Each AccessObject In Application.CurrentDb.TableDefs
                Recordset.AddNew
                Recordset![ObjectName] = AccessObject.Name
                Recordset![ObjectType] = ObjectType
                Recordset.Update
            Next
        End If
    ElseIf ObjectType = acQuery Then
        If stringQueryName <> "" And blnFields Then
            stringSQLText = "DELETE temporaryAccessObjects.ObjectType" & vbCrLf
            stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
            stringSQLText = stringSQLText & "       WHERE (((temporaryAccessObjects.ObjectType)=" & 13 & "));"
            codeDatabase.Execute stringSQLText, dbFailOnError
            Set currentDatabase = CurrentDb
            Dim queryDefinition As DAO.QueryDef
            Set queryDefinition = currentDatabase.QueryDefs(stringQueryName)
            For Each DAOField In queryDefinition.Fields
                Recordset.AddNew
                Recordset![ObjectName] = DAOField.Name
                Recordset![ObjectType] = 13: ObjectType = 13
                Recordset.Update
            Next
                
        Else
            stringSQLText = "DELETE temporaryAccessObjects.ObjectType" & vbCrLf
            stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
            stringSQLText = stringSQLText & "       WHERE (((temporaryAccessObjects.ObjectType)=" & ObjectType & "));"
            codeDatabase.Execute stringSQLText, dbFailOnError
            For Each AccessObject In Application.CurrentDb.QueryDefs
                If Left(AccessObject.Name, 1) <> "~" Then
                    Recordset.AddNew
                    Recordset![ObjectName] = AccessObject.Name
                    Recordset![ObjectType] = ObjectType
                    Recordset.Update
                End If
            Next
        End If
    
    End If
    stringSQLText = "SELECT temporaryAccessObjects.ObjectName" & vbCrLf
    stringSQLText = stringSQLText & "        FROM temporaryAccessObjects" & vbCrLf
    stringSQLText = stringSQLText & " WHERE temporaryAccessObjects.ObjectType =" & ObjectType & ""
    stringSQLText = stringSQLText & "    ORDER BY temporaryAccessObjects.ObjectName" & vbCrLf
    Control.RowSource = stringSQLText
    Control.RowSourceType = "Table/Query"
    Control.Requery
ExitHere:
   Exit Sub

FillListwithObjectNames_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure FillListwithObjectNames of Module General" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select

End Sub

Public Function ProperCase(stringValue As String) As String
    If Len(stringValue & "") = 0 Then Exit Function
    Dim blnLastCharacterSpace As Boolean
    blnLastCharacterSpace = False
    Dim integerCounter As Integer
    Dim stringTemporary As String
    For integerCounter = 1 To Len(stringValue)
        If blnLastCharacterSpace Then
            stringTemporary = stringTemporary & UCase(Mid(stringValue, integerCounter, 1))
        Else
            If integerCounter = 1 Then
                stringTemporary = stringTemporary & UCase(Mid(stringValue, integerCounter, 1))
            Else
                stringTemporary = stringTemporary & Mid(stringValue, integerCounter, 1)
            End If
        End If
        If Mid(stringValue, integerCounter, 1) = " " Then
            blnLastCharacterSpace = True
        Else
            blnLastCharacterSpace = False
        End If
    Next
    ProperCase = stringTemporary
End Function