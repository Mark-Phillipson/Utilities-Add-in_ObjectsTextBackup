Option Compare Database
Option Explicit


Public Function SendEmailMessageNoAttach(strToWhom As String, blnSeeOutlook As Boolean, _
    strTitle As String, strMessage As String)
    Dim intSeeOutlook As Integer
    ' Prevent error screen if user cancels without sending mail.
    On Error Resume Next
    ' Determine if user wants to preview message in Outlook window.
    
    If blnSeeOutlook Then
        intSeeOutlook = True
    Else
        intSeeOutlook = False
    End If

    ' If user wants to directly send item, get recipient's address.
    If intSeeOutlook = False Then
        If IsNull(strToWhom) Then
            strToWhom = InputBox("Enter recipient's e-mail address.")
        End If
    End If
    ' Send object in specified format.
    ' Open Outlook window if intSeeOutlook is True.
    ' Provide Subject title bar and message text.
    DoCmd.SendObject , , , strToWhom, , , strTitle, strMessage, intSeeOutlook

End Function

Public Function SendMailsortRepToJSR()

   DoCmd.OpenForm "frmEmailObject", acNormal, , , acFormEdit, acWindowNormal, "Mailsort Breakdown"

End Function

Public Function SendEmailToAE()
    Dim varTemp As Variant
    Dim f As Form
    DoCmd.OpenForm "frmLiveJobSelection", acNormal, , , acFormEdit, acDialog
    If IsLoaded("frmLiveJobSelection") Then
        Set f = Forms![frmLiveJobSelection]
        varTemp = SendEmailMessageNoAttach("", True, "Ref: " & f![cboLiveJob].Column(2) _
        & " " & f![txtIssue] & " - " & f![cboLiveJob].Column(1), "Customer: " & f![cboLiveJob].Column(3))
    DoCmd.Close acForm, "frmLiveJobSelection", acSaveNo
    End If
    
End Function
Public Sub SendEmailMessage(blnShowMsg As Boolean, strSubject As String, strBody As String, _
    Optional strToWhom As String, Optional strCC As String, _
    Optional AttachmentPath As String)


   Dim olookApp As Object 'Outlook.Application
   Dim olookMsg As Object 'Outlook.MailItem
   Dim olookRecipient As Object 'Outlook.Recipient
   Dim olookAttach As Object 'Outlook.Attachment
   Dim blnMultiSelection As Boolean
   Dim kounter As Integer, strTemp As String
   Dim strFilesSelected(30) As String
   Dim strWhomTemp As String
   Dim strCCTemp As String
   On Error GoTo HandleErr
   Const olMailItem = 0
   Const olTo = 1
   Const olCC = 2
   Const olImportanceLow = 0
   
    If InStr(1, AttachmentPath, Chr(9), vbTextCompare) > 0 Then
        blnMultiSelection = True
    Else
        blnMultiSelection = False
    End If
   ' create the Outlook session.
   Set olookApp = CreateObject("Outlook.Application")

   ' create the message.
   Set olookMsg = olookApp.CreateItem(olMailItem)

   With olookMsg
      ' add the To recipient(s) to the message.
          
      If Not IsMissing(strToWhom) Then
        If Len(strToWhom) > 0 Then
          If InStr(strToWhom, "|") > 0 Then
            Do
                strWhomTemp = Left(strToWhom, InStr(strToWhom, "|") - 1)
                Set olookRecipient = .Recipients.Add(strWhomTemp)
                olookRecipient.Type = olTo
                strToWhom = Mid(strToWhom, InStr(strToWhom, "|") + 1)
                If InStr(strToWhom, "|") = 0 Then Exit Do
            Loop
            Set olookRecipient = .Recipients.Add(strToWhom)
            olookRecipient.Type = olTo
          Else
            Set olookRecipient = .Recipients.Add(strToWhom)
            olookRecipient.Type = olTo
          End If
        End If
      End If

      ' add the CC recipient(s) to the message.
      If Not IsMissing(strCC) Then
        If Len(strCC) > 0 Then
          If InStr(strCC, "|") > 0 Then
            Do
                strCCTemp = Left(strCC, InStr(strCC, "|") - 1)
                Set olookRecipient = .Recipients.Add(strCCTemp)
                olookRecipient.Type = olCC
                strCC = Mid(strCC, InStr(strCC, "|") + 1)
                If InStr(strCC, "|") = 0 Then Exit Do
            Loop
            Set olookRecipient = .Recipients.Add(strCC)
            olookRecipient.Type = olCC
          Else
          Set olookRecipient = .Recipients.Add(strCC)
          olookRecipient.Type = olCC
          End If
        End If
      End If
      ' set the Subject, Body, and Importance of the message.
      .Subject = strSubject
      .Body = strBody
      .IMPORTANCE = olImportanceLow 'Low importance

      ' add attachments to the message.
      If Len(AttachmentPath) > 0 Then
         If blnMultiSelection Then
                strFilesSelected(1) = Left(AttachmentPath, InStr(1, AttachmentPath, Chr(9), vbTextCompare) - 1)
                strTemp = Mid(AttachmentPath, InStr(1, AttachmentPath, Chr(9), vbTextCompare) + 1)
                Set olookAttach = .Attachments.Add(strFilesSelected(1))
                For kounter = 2 To 30
                    If InStr(1, strTemp, Chr(9), vbTextCompare) = 0 Then
                        strFilesSelected(kounter) = strTemp
                        Set olookAttach = .Attachments.Add(strFilesSelected(kounter))
                        Exit For
                    Else
                        strFilesSelected(kounter) = Mid(strTemp, 1, InStr(1, strTemp, Chr(9), vbTextCompare) - 1)
                    End If
                    Set olookAttach = .Attachments.Add(strFilesSelected(kounter))
                    strTemp = Mid(strTemp, InStr(1, strTemp, Chr(9), vbTextCompare) + 1)
                Next
         Else
            Set olookAttach = .Attachments.Add(AttachmentPath)
         End If
      End If

      ' resolve each Recipient's name
      For Each olookRecipient In .Recipients
         olookRecipient.Resolve
         If Not olookRecipient.Resolve Then
            olookMsg.Display   ' display any names that can't be resolved
         End If
      Next
        
      If blnShowMsg Then
          olookMsg.Display
      Else
          .Send
      End If
         
        
    End With
    Set olookMsg = Nothing
    Set olookApp = Nothing
      
ExitHere:
    Exit Sub

HandleErr:
    Select Case Err.Number
    Case 287 'Application-defined or object-defined error
    'When user says no to allow access to Outlook
    MsgBox "Email cannot be created - permission denied by user????", vbExclamation, "Warning"
    Resume ExitHere
    Case Else
        MsgBox "Unexpected Error Please inform support " & Err.Number & ": " & Err.Description, vbCritical, "basSendEmail.SendEmailMessage" 'ErrorHandler:$$N=basSendEmail.SendEmailMessage
        Resume ExitHere
    End Select
Resume 'Debug Only
   End Sub
Sub ShowFreeSpace(drvPath)
    Dim fs, d, s As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " & UCase(drvPath) & " - "
    s = s & d.VolumeName & vbCrLf
    s = s & "Free Space: " & FormatNumber(d.FreeSpace / 1024, 0)
    s = s & " Kbytes"
    MsgBox s
    
End Sub
Sub ShowDriveList()
    Dim fs, d, dc, s, N, Remote, folder
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = "Temp"
    If fs.FolderExists(folder) Then
        MsgBox folder & " exists"
    Else
        MsgBox folder & " does not exists"
    End If
End Sub