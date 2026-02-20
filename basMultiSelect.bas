Option Compare Database
    
Public Sub PrepareMultiSelect(strSQLtext As String, ListboxToSelect As ListBox, Optional intColumnPosition As Integer = 0, Optional blnTableInAddin As Boolean = False)
    'to prepare the multi-select supply a SQL text to list 1 text item and the listbox on the source form
    'Make sure the column name is called value and any example call is as follows:
    '    strSQLText = "SELECT qryStructure.[Column Name] as [Value]" & vbCrLf
    'strSQLText = strSQLText & "        FROM qryStructure" & vbCrLf
    'strSQLText = strSQLText & "    ORDER BY qryStructure.[Column Name];"
    'PrepareMultiSelect strSQLText, Me.lstStructure
    Dim m As Integer 'Aug 11
    Dim k As Integer 'Aug 11
    
    Dim rstRecordSet As DAO.Recordset
    Dim database As DAO.database
    If blnTableInAddin Then
        Set database = codeDB
    Else
        Set database = CurrentDb
    End If

    Set rstRecordSet = database.OpenRecordset(strSQLtext, dbReadOnly)
    DoCmd.Close acForm, "frmFormMultiPic", acSaveYes
    Set frmFormMultiPic = New Form_FormMultiPic
    frmFormMultiPic.SetupSourceListBox rstRecordSet
    'Setup currently selected
    frmFormMultiPic.ListboxSelected.RowSource = ""
    For k = 0 To ListboxToSelect.ListCount - 1
        If ListboxToSelect.Selected(k) = True Then
            frmFormMultiPic.ListboxSelected.AddItem (ListboxToSelect.Column(intColumnPosition, k))
            frmFormMultiPic.ListboxSource.RemoveItem (ListboxToSelect.Column(intColumnPosition, k))
        End If
    Next
    frmFormMultiPic.TextBoxSourceCount = frmFormMultiPic.ListboxSource.ListCount
    frmFormMultiPic.SetFocus
    Do
        If Not IsLoaded("FormMultiPic") Then Exit Do
        If frmFormMultiPic.Visible = False Then Exit Do
        DoEvents
    Loop
        
        
        
        
    If IsLoaded("FormMultiPic") Then
        
        For m = 0 To ListboxToSelect.ListCount - 1
            ListboxToSelect.Selected(m) = False
        Next
        
        For m = 0 To ListboxToSelect.ListCount - 1
        
            
            For k = 0 To frmFormMultiPic.ListboxSelected.ListCount
                If frmFormMultiPic.ListboxSelected.ItemData(k) = ListboxToSelect.Column(intColumnPosition, m) Then
                    ListboxToSelect.Selected(m) = True

                    
                End If
            Next
        Next
    End If
    If ListboxToSelect.Enabled Then
        ListboxToSelect.SetFocus
        
    End If

End Sub