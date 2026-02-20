 Option Compare Database

'retrieve IDs of all CommandBarControl objects by enumerating every CommandBar
'object in the CommandBars collection
Sub DumpBars()
PrintLine "Command bar controls", 1
Application.TempVars.Add "TemporaryDump", "-"
Dim bar As CommandBar
For Each bar In Application.VBE.CommandBars
DumpBar bar, 2

'Application.VBE.CommandBars("Tools").Controls.Add
Next
     Debug.Print Application.TempVars![TemporaryDump]
End Sub

'retrieve IDs of all CommandBarControl objects in a CommandBar object
Sub DumpBar(bar As CommandBar, level As Integer)
PrintLine "CommandBar Name: " & bar.Name, level
Dim ctl As commandBarControl
For Each ctl In bar.Controls
DumpControl ctl, level + 1
Next

End Sub

'retrieve the ID of a CommandBarControl object. If it is a CommandBarPopup object,
'which could serve as a control container, try to enumerate all sub controls.
Sub DumpControl(ctl As commandBarControl, level As Integer)
PrintLine vbTab & "Control Caption: " & ctl.Caption & vbTab & "ID: " & ctl.ID, level
Select Case ctl.Type
'the following control type could contain sub controls
Case msoControlPopup, msoControlGraphicPopup, msoControlButtonPopup, msoControlSplitButtonPopup, msoControlSplitButtonMRUPopup
Dim sctl As commandBarControl, ctlPopup As CommandBarPopup
Set ctlPopup = ctl
For Each sctl In ctlPopup.Controls
DumpControl sctl, level + 1
Next
Case Else
'no sub controls
End Select
End Sub

'Word only output function. You should modify this if you are working with
'other Office products.
Sub PrintLine(line As String, level As Integer)
'Selection.Paragraphs(1).OutlineLevel = level
If InStr(line, "Comment") = 0 Then Exit Sub
Application.TempVars![TemporaryDump] = Application.TempVars![TemporaryDump] & Space(level * 4) & line & vbNewLine
'Selection.Collapse wdCollapseEnd
End Sub