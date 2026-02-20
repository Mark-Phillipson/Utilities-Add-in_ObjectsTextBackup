Option Compare Database
Option Explicit

Public Function DisplayASCIICharacters(variantvalue As Variant) As Variant
    Dim stringValue As String
    Dim longCounter As Long
    Dim stringTemporary As String
    If IsNull(variantvalue) Then
        Exit Function
    End If
    stringValue = variantvalue
    For longCounter = 1 To Len(stringValue)
        stringTemporary = stringTemporary & longCounter & ": " & Asc(Mid(stringValue, longCounter, 1)) & " "
    Next
    DisplayASCIICharacters = stringTemporary
End Function

Public Function ShownASCIICharacters()
    
    On Error GoTo ShownASCIICharacters_Error
    Dim stringValue As String
    stringValue = Nz(Screen.ActiveDatasheet.activeControl.Value, "")
    If Len(stringValue & "") = 0 Then
        Exit Function
    End If
    
    MsgBox DisplayASCIICharacters(stringValue), vbInformation, "ASCII Characters (Position: ASCII)"

ExitHere:
   Exit Function

ShownASCIICharacters_Error:
    Select Case Err.Number
    Case 2484  'There is no active datasheet.
        Resume ExitHere
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure ShownASCIICharacters of Module FixControlCharacters" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Function