Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Compare Database
Option Explicit

Private WithEvents ListFormsMenuCommand As CommandBarButton
Attribute ListFormsMenuCommand.VB_VarHelpID = -1
Private WithEvents ListReportsMenuCommand As CommandBarButton
Attribute ListReportsMenuCommand.VB_VarHelpID = -1
Private WithEvents ListTablesMenuCommand As CommandBarButton
Attribute ListTablesMenuCommand.VB_VarHelpID = -1
Private WithEvents ListQueriesMenuCommand As CommandBarButton
Attribute ListQueriesMenuCommand.VB_VarHelpID = -1

Private Sub ListFormsMenuCommand_Click(ByVal Ctrl As CommandBarButton, blnCancel As Boolean)
    InsertTextIntoModule GetObjectName(acForm)
End Sub

Private Sub ListReportsMenuCommand_Click(ByVal Ctrl As CommandBarButton, blnCancel As Boolean)
    InsertTextIntoModule GetObjectName(acReport)
End Sub

Private Sub ListTablesMenuCommand_Click(ByVal Ctrl As CommandBarButton, blnCancel As Boolean)
    InsertTextIntoModule GetObjectName(acTable)
End Sub

Private Sub ListQueriesMenuCommand_Click(ByVal Ctrl As CommandBarButton, blnCancel As Boolean)
    InsertTextIntoModule GetObjectName(acQuery)
End Sub

'Public Property Get ListFormsMenuCommandProperty() As CommandBarButton
'    Set ListFormsMenuCommandProperty = ListFormsMenuCommand
'
'End Property

Public Property Let ListFormsMenuCommandProperty(ByVal NewListFormsMenuCommand As CommandBarButton)
    Set ListFormsMenuCommand = NewListFormsMenuCommand
End Property

Private Function GetObjectName(AccessObject As AcObjectType) As String
    DoCmd.Close acForm, "Generic Access Object Picker", acSaveYes
    DoCmd.OpenForm "Generic Access Object Picker", acNormal, , , acFormEdit, acDialog, AccessObject
    If IsLoaded("Generic Access Object Picker") Then
        GetObjectName = Forms![Generic Access Object Picker].ResultsListBox
    End If

End Function

Public Sub InsertTextIntoModule(stringValue As String)
    With Application.VBE.ActiveCodePane.CodeModule
        .InsertLines 1, stringValue
    End With

End Sub