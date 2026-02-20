Option Compare Database
Option Explicit
Global longID As Long
Global stringSelectedItem As String

Public Function DoGenericSearch(stringSelectText As String, stringWhereClause As String, stringOrderByClause As String, _
    integerColumns As Integer, stringColumnWidths As String, Optional stringRowSourceType As String = "Table/Query") As Variant
    Dim Form As [Form_Generic Search]
    On Error GoTo DoGenericSearch_Error

    On Error Resume Next
    DoCmd.Close acForm, "Generic Search", acSaveYes
    On Error GoTo DoGenericSearch_Error
    Set Form = New [Form_Generic Search]
    Form.Visible = False
    
    Form.InitialSetup stringSelectText, stringWhereClause, stringOrderByClause, integerColumns, stringColumnWidths, stringRowSourceType
    
    Form.Visible = True
    Form.SetFocus
    Do Until Not IsLoaded("Generic Search")
        DoEvents
    Loop
    If stringRowSourceType = "Table/Query" Then
        DoGenericSearch = longID
    Else
        DoGenericSearch = stringSelectedItem
    End If

ExitHere:
   Exit Function

DoGenericSearch_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure DoGenericSearch of Module Module1" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function

Public Sub TestGenericSearch()
    Dim stringSQLText As String
    stringSQLText = "SELECT Customers.ID" & vbCrLf
    stringSQLText = stringSQLText & "           , Customers.CompanyName" & vbCrLf
    stringSQLText = stringSQLText & "           , Customers.EmailAddress" & vbCrLf
    stringSQLText = stringSQLText & "           , Customers.ContactName" & vbCrLf
    stringSQLText = stringSQLText & "        FROM Customers" & vbCrLf
    Dim stringSelectText As String
    stringSelectText = stringSQLText
    Dim stringWhereClause As String
    stringWhereClause = "       WHERE ((([Customers]![CompanyName] & "" "" & [Customers]![EmailAddress] & "" "" & [Customers]![ContactName]) Like '*'))" & vbCrLf
    Dim stringOrderByClause As String
    stringOrderByClause = "    ORDER BY Customers.ContactName;"
    Debug.Print DoGenericSearch(stringSelectText, stringWhereClause, stringOrderByClause, 4, "0cm;5cm;5cm;5cm")
End Sub
Public Sub TestGeneralSearchValueList()
    Debug.Print DoGenericSearch("Item A;Item B; Item C", "", "", 1, "3cm", "Value List")
End Sub

Public Function ConvertDelimitedToArray(strDelimiter As String _
    , strValue As String) As Variant
    Dim intDelimiters As Integer 'Added 14/10/2005
    Dim strTemp As String
    Dim intPosn As Integer
    Dim k As Integer
    Dim varArray() As Variant
    'ReDim Preserve X(10, 10, 15)
    'Count number of delimiters
    intDelimiters = 0
    For k = 1 To Len(strValue)
        If Mid(strValue, k, Len(strDelimiter)) = strDelimiter Then
            intDelimiters = intDelimiters + 1
        End If
    Next
    intDelimiters = intDelimiters + 1
    ReDim varArray(intDelimiters) As Variant
    k = -1
    If Len(strValue & "") > 0 And InStr(strValue, strDelimiter) > 0 Then
        strTemp = strValue
        Do Until InStr(strTemp, strDelimiter) = 0
            intPosn = InStr(strTemp, strDelimiter)
            k = k + 1
            
            varArray(k) = Left(strTemp, intPosn - 1)
            strTemp = Mid(strTemp, intPosn + Len(strDelimiter))
        Loop
        varArray(k + 1) = strTemp
        
    End If
    ConvertDelimitedToArray = varArray
End Function