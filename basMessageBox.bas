Option Compare Database
Option Explicit

Public Enum WhichButton
    mbButton1 = 1
    mbButton2 = 2
    mbButton3 = 3
    mbDefault = 4
End Enum

Public Enum WhatIcon
    iconNone
    iconCritical
    iconExclamation
    iconInformation
    iconQuestion
End Enum


Public Enum btIcon
    btNone
    btInformation
    btWarning
    btCritical
End Enum


Public Function MessageBox(strCaption As String, strMessage As String, strTitle As String _
, strButton1Caption As String, strButton2Caption As String, strButton3Caption As String _
, strDefaultCaption As String, Icon As WhatIcon, Optional intDefaultButton As Integer) As WhichButton
    'Example call
    'If MessageBox("Form Cap", "message text", "The Title", "But1", "But2", "But3", "Def", iconNone) = mbDefault Then
    
    
    Dim tv As TaggedValues
    On Error GoTo HandleErr
    MessageBox = 0
    Set tv = New TaggedValues
    
    tv.Add "Caption", strCaption
   ' tv.Add "Message", strMessage
   TempVars("Message") = strMessage
    tv.Add "Title", strTitle
    tv.Add "Button1", strButton1Caption
    tv.Add "Button2", strButton2Caption
    tv.Add "Button3", strButton3Caption
    tv.Add "Default", strDefaultCaption
    tv.Add "Icon", str(Icon)
    tv.Add "DefaultButton", str(intDefaultButton)
    DoCmd.Close acForm, "frmMessageBox", acSaveYes
   
    DoCmd.OpenForm "frmMessageBox", acNormal, , , acFormEdit, acDialog, tv.Text
    
    MessageBox = Forms![frmMessageBox].txtResult
    
    Set tv = Nothing



ExitHere:
    Exit Function

HandleErr:
    Select Case Err.Number
    'Case  '
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "basMessageBox.MessageBox"   'ErrorHandler:$$N=basMessageBox.MessageBox
        Resume ExitHere
    End Select
    Resume 'Debug Only
End Function

Public Function ShowBalloonTooltip(strHeading As String, strMessage As String, lngIcon As btIcon)
'    Set bt = New BalloonTooltip
'    With bt
'        .Heading = strHeading
'        .message = strMessage
'        .Icon = lngIcon
'        .Show
'        .Show
'    End With
    MsgBox strMessage, vbCritical, strHeading
End Function
'
'Public Function HideIcon()
'    If Not bt Is Nothing Then
'        With bt
'            .Hide
'        End With
'    End If
'End Function