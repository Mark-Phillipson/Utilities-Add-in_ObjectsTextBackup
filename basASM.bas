Option Compare Database
Option Explicit
Public Type BridgeDetails
    lngIndiciaID As Long
    strCageName As String
    strCroydonFlatCageCode As String
    strCroydonPacketCageCode As String
    strDominoSplit As String
End Type
Public Enum mpFormat
    mpLetter = 1
    mpFlat = 2
    mpPacket = 3
    mpMBag = 4
End Enum

Public Function AddFieldGenericNoError(TableName As String, FieldName As String, intDataType As Integer, _
    Optional intFieldSize As Integer = 50)
Dim MyDatabase As DAO.database
Dim MyTableDef As DAO.tableDef
Dim MyField As DAO.field, i As Integer
Dim answer As Integer

On Error GoTo AddFieldHandler
Set MyDatabase = CurrentDb
MyDatabase.TableDefs.Refresh  ' Refresh possibly changed collection.
For i = 0 To MyDatabase.TableDefs.Count - 1
    If MyDatabase.TableDefs(i).Name = TableName Then Exit For
Next
Set MyTableDef = MyDatabase.TableDefs(i)
' Create new Field object.
If IsNull(FieldName) Then FieldName = "ZZ"
Set MyField = MyTableDef.CreateField(FieldName, intDataType)
' Set another property of MyField.
If intDataType = dbText Then
    MyField.Size = intFieldSize
    MyField.AllowZeroLength = True
End If

MyTableDef.Fields.Refresh
' Save Field definition by appending it to Fields collection.
MyTableDef.Fields.Append MyField
MyTableDef.Fields.Refresh
AddFieldGenericNoError = True

Exit Function
AddFieldHandler:

Select Case Err
Case Is = 3191
    'MsgBox "Add Field Error.  Can't define field more than once - " & FieldName, vbInformation
    AddFieldGenericNoError = False
    Exit Function
Case Is = 3211
    'MsgBox Err.DESCRIPTION
    'answer = MsgBox("Error No: " & Err.number & " has occured " & Err.Description, vbRetryCancel + vbDefaultButton1)
    'If answer = vbRetry Then
    '    MyDatabase.TableDefs.Refresh  ' Refresh possibly changed collection.
    '    Resume
    'Else
        Exit Function
    'End If
    
Case Else
    MsgBox "Error No: " & Err.Number & " has occured " & Err.Description
    Stop
    Resume

End Select

End Function