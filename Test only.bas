Option Compare Database
Option Explicit



Public Sub PhilPropertiesTable()
    Dim Form As Access.Form
Dim Report As Access.Report
Dim TextBox As TextBox

    On Error GoTo PhilPropertiesTable_Error

    DoCmd.OpenForm "frmTesting", acDesign
'DoCmd.OpenReport "rptCompare2Tables", acViewDesign

    Set Form = Forms![frmTesting]
'Set Report = Reports![rptCompare2Tables]
'Set TextBox = Report![txtTbl2_Field_Name]
    Dim currentDatabase As DAO.database
        Dim Recordset As DAO.Recordset
        Dim stringSQLText As String
        Set currentDatabase = CurrentDb
        Set Recordset = currentDatabase.OpenRecordset("Object Properties")
        Dim Property As Property
        For Each Property In Form.Properties
            Recordset.AddNew
            Recordset![PropertyName] = Property.Name
            Recordset![FormApplicable] = True
            Recordset![Type] = Property.Type
            Recordset![SourceObject] = "Form"
            On Error Resume Next
            Recordset![ExampleValue] = Property.Value

            Recordset.Update
            On Error GoTo PhilPropertiesTable_Error
'            If property.Name = Nz(DLookup("[PropertyName]", "[Object Properties]", "[PropertyName] = '" & property.Name & "'"), "") Then
'                stringSQLText = "UPDATE [Object Properties] SET [Object Properties].TextBoxApplicable = True" & vbCrLf
'                stringSQLText = stringSQLText & "       WHERE ((([Object Properties].PropertyName)='" & property.Name & "'));"
'                CurrentDatabase.Execute stringSQLText, dbFailOnError
'            End If
            
        Next
        Recordset.Close

ExitHere:
   Exit Sub

PhilPropertiesTable_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure PhilPropertiesTable of Module Test only" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only

            
End Sub

Public Sub PhilPropertiesfromReport()
Dim Report As Access.Report
Dim TextBox As TextBox


'    DoCmd.OpenForm "frmTesting", acDesign
    On Error GoTo PhilPropertiesfromReport_Error

DoCmd.OpenReport "rptCompare2Tables", acViewDesign

Set Report = Reports![rptCompare2Tables]
'Set TextBox = Report![txtTbl2_Field_Name]
    Dim currentDatabase As DAO.database
        Dim Recordset As DAO.Recordset
        Dim stringSQLText As String
        Set currentDatabase = CurrentDb
        Set Recordset = currentDatabase.OpenRecordset("Object Properties")
        Dim Property As Property
        For Each Property In Report.Properties
            Recordset.AddNew
            Recordset![PropertyName] = Property.Name
            Recordset![ReportApplicable] = True
            Recordset![Type] = Property.Type
            Recordset![SourceObject] = "Report"
            On Error Resume Next
            Recordset![ExampleValue] = Property.Value
            Recordset.Update
            On Error GoTo PhilPropertiesfromReport_Error
            If Property.Name = Nz(DLookup("[PropertyName]", "[Object Properties]", "[PropertyName] = '" & Property.Name & "'"), "") Then
                stringSQLText = "UPDATE [Object Properties] SET [Object Properties].ReportApplicable = True" & vbCrLf
                stringSQLText = stringSQLText & "       WHERE ((([Object Properties].PropertyName)='" & Property.Name & "'));"
                currentDatabase.Execute stringSQLText, dbFailOnError
            End If
            
        Next
        Recordset.Close

ExitHere:
   Exit Sub

PhilPropertiesfromReport_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure PhilPropertiesfromReport of Module Test only" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Sub

Public Sub PhilControlProperties(stringControlType As String, stringControlName As String)
Dim Form As Access.Form
Dim Control As Access.Control


    DoCmd.OpenForm "frmTesting", acDesign

    On Error GoTo PhilControlProperties_Error

'DoCmd.OpenForm "rptCompare2Tables", acViewDesign

Set Form = Forms![frmTesting]
Set Control = Form.Controls(stringControlName)
    Dim currentDatabase As DAO.database
        Dim Recordset As DAO.Recordset
        Dim stringSQLText As String
        Set currentDatabase = CurrentDb
        Set Recordset = currentDatabase.OpenRecordset("Object Properties")
        Dim Property As Property
        For Each Property In Control.Properties
            Recordset.AddNew
            Recordset![PropertyName] = Property.Name
            Recordset.Fields(stringControlType & "Applicable") = True
            Recordset![Type] = Property.Type
            Recordset![SourceObject] = stringControlType
            On Error Resume Next
            Recordset![ExampleValue] = Property.Value
            Recordset.Update
            On Error GoTo PhilControlProperties_Error
            If Property.Name = Nz(DLookup("[PropertyName]", "[Object Properties]", "[PropertyName] = '" & Property.Name & "'"), "") Then
                stringSQLText = "UPDATE [Object Properties] SET [Object Properties]." & stringControlType & "Applicable = True" & vbCrLf
                stringSQLText = stringSQLText & "       WHERE ((([Object Properties].PropertyName)='" & Property.Name & "'));"
                currentDatabase.Execute stringSQLText, dbFailOnError
            End If
            
        Next
        Recordset.Close

ExitHere:
   Exit Sub

PhilControlProperties_Error:
    Select Case Err.Number
    'Case '
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Service Desk." _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure PhilControlProperties of Module Test only" _
            , vbCritical
        Resume ExitHere
    End Select
    Resume 'For Debug Only


End Sub
Public Function Test()
    Dim queryDefinition As DAO.QueryDef
    Set queryDefinition = CurrentDb.QueryDefs("")
    Dim DAOField As DAO.field
    Dim integerCounter As Integer
    For Each DAOField In queryDefinition.Fields
        If DAOField.Required Then
            MsgBox "this field is required " & DAOField.Name, vbInformation, "Required"
        End If
    Next
End Function