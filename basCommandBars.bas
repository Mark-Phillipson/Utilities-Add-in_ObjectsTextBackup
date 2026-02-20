Option Explicit
Dim Command As CommandBarClass
Dim commandBarControl As commandBarControl
'Private Const conMenuIDTools = 30007
Private Const conMenuName = "&VBA Code Writer"


Public Sub SetupVBEMenu()
    Set commandBarControl = AddMenuCommand(Application.VBE.CommandBars("Menu Bar"), _
               Application, "Access List Forms", "Access List Forms", 0, "Access List Forms", "")
    Set Command = New CommandBarClass
    Command.ListFormsMenuCommandProperty = commandBarControl
End Sub

Public Function AddMenuCommand(ByRef cbrMenu As CommandBar, ByRef AddInInst As Object, Optional ByRef strProp As String = "", Optional ByRef strValue As String = "", Optional ByRef lngFaceid As Integer = 0, Optional ByRef conMenuName As String = "&VBA Code Writer", Optional stringOnAction As String = "") As commandBarControl


        Dim cbcMsgBox As commandBarControl
        Dim cbcFormat As commandBarControl
        Dim c As commandBarControl
        Dim cbr As CommandBar
        Dim cbc As commandBarControl
        Dim blnValue As Boolean
        On Error GoTo HandleErr
        Set AddMenuCommand = Nothing
        If cbrMenu.Name = "Tools" Then
            ' Get a pointer to the Tools menu
            'Set cbcTools = cbrMenu.FindControl( _
            ''Name:="Code Window", Recursive:=False)

            ' If we found the Tools menu then add
            ' a new menu command (but only if it doesn't
            ' already exist!)
            If Not cbrMenu Is Nothing Then

                ' Try to find the command based on its tag
                Set cbcMsgBox = cbrMenu.FindControl(tag:=conMenuName, Recursive:=False)

                ' If we didn't find it, add a new command
                If cbcMsgBox Is Nothing Then
                    Set cbcMsgBox = cbrMenu.Controls.Add(Type:=MsoControlType.msoControlButton)
                    With cbcMsgBox
                        .Caption = conMenuName
                        'UPGRADE_WARNING: Couldn't resolve default property of object cbcMsgBox.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .style = MsoButtonStyle.msoButtonIconAndCaption
                        .tag = conMenuName
                        'UPGRADE_WARNING: Couldn't resolve default property of object cbcMsgBox.FaceId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        .FaceId = lngFaceid

                        ' This enables demand loading
                        .OnAction = stringOnAction
                        .BeginGroup = True
                        .TooltipText = strValue
                        .DescriptionText = strValue
                        .Visible = True
                    End With
                End If

                ' Return pointer to menu command
                Set AddMenuCommand = cbcMsgBox
                'Debug.Print cbcMsgBox.Caption
            End If
        Else
            Set cbr = cbrMenu
            Set cbc = cbr.FindControl(tag:=strValue)
            If strValue = "True" Then
                blnValue = True
            Else
                blnValue = False
            End If
            If cbc Is Nothing Then
                Set cbc = cbr.Controls.Add(Type:=MsoControlType.msoControlButton)
                With cbc
                    .Caption = strValue
                    .tag = .Caption
                    '.OnAction = "=MercVBACodeWriter.ToggleProperty(" & Chr(34) & strProp & Chr(34) & ", True)"
                    .BeginGroup = False
                    .TooltipText = .Caption
                    .DescriptionText = .Caption
                    .Visible = True
                    'UPGRADE_WARNING: Couldn't resolve default property of object cbc.FaceId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .FaceId = lngFaceid
                    'UPGRADE_WARNING: Couldn't resolve default property of object cbc.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .style = MsoButtonStyle.msoButtonIconAndCaption
                    .OnAction = stringOnAction
                    .Parameter = "Test 12345"

                End With
            End If
            Set cbcFormat = cbr.FindControl(tag:=strValue)
            Set AddMenuCommand = cbcFormat

        End If



ExitHere:
        Exit Function

        ' Error handling block added by Error Handler Add-In. DO NOT EDIT this block of code.
        ' Automatic error
HandleErr:
        Select Case Err.Number
            'Case # '
            ' MsgBox "", vbExclamation, ""
            Case Else
                'MsgBox("Error " & Err.Number & ": " & Err.Description, MsgBoxStyle.Critical, "Unexpected Error in basCommandBars.AddMenuCommand") 'ErrorHandler:$$N=basCommandBars.AddMenuCommand
                Resume ExitHere
        End Select
        Resume
    End Function



Public Function DisplayMessage(stringMessage As String)
    MsgBox stringMessage, vbInformation, "Display"
End Function