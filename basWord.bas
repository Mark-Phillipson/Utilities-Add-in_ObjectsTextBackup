Option Compare Database
Option Explicit

Enum pgeOrientation
    pgePortrait = 1
    pgeLandsape = 2
End Enum



Function CreateTableFromRecordset( _
 rngAny As Object, _
 rstAny As ADODB.Recordset, _
 Optional fIncludeFieldNames As Boolean = False) _
 As Object

    Dim objTable As Object
    Dim fldAny As ADODB.field
    Dim varData As Variant
    Dim strBookmark As String
    Dim cField As Long
    Const wdSeparateByTabs = 1
    Const wdAutoFitContent = 1

    ' Get the data from the recordset
    varData = rstAny.GetString()
    rstAny.MoveLast
    rstAny.MoveFirst
    'Debug.Print rstAny.RecordCount
    ' Create the table
    With rngAny
    
        ' Creating the basic table is easy,
        ' just insert the tab-delimted text
        ' add convert it to a table
        .InsertAfter varData
        Set objTable = .ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=rstAny.Fields.Count, _
        NumRows:=rstAny.RecordCount, AutoFitBehavior:=wdAutoFitContent)
        With objTable
            .style = "Table Grid"
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastRow = True
            .ApplyStyleFirstColumn = True
            .ApplyStyleLastColumn = True
        End With
        
        
        ' Field names are more work since
        ' you must do them one at a time
        If fIncludeFieldNames Then
            With objTable
            
                ' Add a new row on top and make it a heading
                .Rows.Add(.Rows(1)).HeadingFormat = True
                
                ' Iterate through the fields and add their
                ' names to the heading row
                For Each fldAny In rstAny.Fields
                    cField = cField + 1
                    .Cell(1, cField).Range.Text = fldAny.Name
                Next
            End With
        End If
    End With
    Set CreateTableFromRecordset = objTable
End Function

Function PrintRstWithWord(strSQLtext As String, strTitle As String, PageOrientation As pgeOrientation)
    Dim strTemplate As String 'Added 18/02/2004
    Dim cField As Long 'Added 17/02/2004
    Dim objWord As Object 'Word.Application
    Dim rst As New ADODB.Recordset
    Dim c As ADODB.connection
    Dim fldAny As ADODB.field
    Const wdTableFormatGrid8 = 23
    Const wdTexture50Percent = 500
    Const wdColorAutomatic = -16777216
    Const wdLineStyleSingle = 1
    Const wdLineWidth075pt = 6
    Const wdColorDarkBlue = 8388608
    Const wdAutoFitWindow = 2
    Const wdAlignParagraphCenter = 1
    Const wdAlignParagraphLeft = 0

    'Set the provider name
   On Error GoTo PrintRstWithWord_Error
   
    Set c = CodeProject.connection
    'Open a recordset with a keyset cursor
    
    ' Launch Word and load the invoice template
'    If PageOrientation = pgeLandsape Then
'        strTemplate = "\\carrieray\carrier\Mercury Documents\Word\Landscape Report Template.dot"
'    ElseIf PageOrientation = pgePortrait Then
'        strTemplate = "\\carrieray\carrier\Mercury Documents\Word\Portrait Report Template.dot"
'    End If
    
    Set objWord = CreateObject("Word.Application")
    objWord.Documents.Add _
    strTemplate
    objWord.Visible = False

    ' Add header information using predefined bookmarks
    With objWord.ActiveDocument.Bookmarks
        .Item("Title").Range.Text = strTitle
    End With
    
    ' Get details from database and create a table
    ' in the document
    
    rst.Open strSQLtext, c, adOpenKeyset, adLockPessimistic
    With CreateTableFromRecordset( _
     objWord.ActiveDocument.Bookmarks("Table").Range, rst, True)
     
        ' Apply formatting
        .AutoFormat wdTableFormatGrid8
        ' Add rows for subtotal, freight, total
            With .Rows.Add
                For Each fldAny In rst.Fields
                    cField = cField + 1
                    .Cells(cField).Range.Select
                    Select Case fldAny.Type
                    Case adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt
                        objWord.Selection.InsertFormula Formula:="=SUM(ABOVE)", NumberFormat:="0"
                        objWord.Selection.Font.Bold = True
                    Case Else
                        With objWord.Selection.Cells.Shading
                            .Texture = wdTexture50Percent
                            .ForegroundPatternColor = wdColorAutomatic
                            .BackgroundPatternColor = wdColorAutomatic
                        End With
                        With objWord.Options
                            .DefaultBorderLineStyle = wdLineStyleSingle
                            .DefaultBorderLineWidth = wdLineWidth075pt
                            .DefaultBorderColor = wdColorDarkBlue
                        End With
                    End Select
                Next
            End With
        
        .AutoFitBehavior wdAutoFitWindow
        
        ' Fix up paragraph alignment
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        .Columns(1).Select
        objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        objWord.Selection.MoveDown
    End With
    objWord.Visible = True

    ' We're done
    Set objWord = Nothing

ExitHere:
   Exit Function

PrintRstWithWord_Error:

    Select Case Err.Number
    Case Else
        MsgBox "An Unexpected Error has occured please inform IT Support Error " & Err.Number & " " & Err.Description & " in procedure PrintRstWithWord of Module basWord", vbCritical, "UtilitesAdd-In"
        Resume ExitHere
    End Select
    'Debug Only
    Resume
End Function


Public Function Test()
Dim strSQLtext As String
  strSQLtext = "SELECT tmpObjects.[Object name]" & vbCrLf
  strSQLtext = strSQLtext & "           , tmpObjects.id" & vbCrLf
  strSQLtext = strSQLtext & "           , tmpObjects.[Object type]" & vbCrLf
  strSQLtext = strSQLtext & "        FROM tmpObjects;"
    
    PrintRstWithWord strSQLtext, "CURRENT TEMP OBJECTS LIST", pgeLandsape
    
End Function

Public Function FormatRecordsetToHTML(rst1 As Variant) As String
Dim strTable As String  'Table as a string
Dim index As Integer    'Index counter

    strTable = "<table border=1 width=500>"     'create table open
    rst1.MoveFirst
    
    'create headings of the recordset using
    strTable = strTable & "<tr>"
    For index = 0 To rst1.Fields.Count - 1
        strTable = strTable & "<td bgcolor=Blue><font color='white'>"
        strTable = strTable & rst1.Fields.Item(index).Name
        strTable = strTable & "</font></td>"
    Next
    strTable = strTable & "</tr>"
    While (Not rst1.EOF)              'loop until we reach the end of the recordset
        strTable = strTable & "<tr>"
        For index = 0 To rst1.Fields.Count - 1
            strTable = strTable & "<td>"
            strTable = strTable & rst1(rst1.Fields.Item(index).Name).Value       'writes the id to the screen
            strTable = strTable & "<br>"                     'moves to the next line on the page
            strTable = strTable & "</td>"
        Next
        strTable = strTable & "</tr>"
        rst1.MoveNext                     'moves to the next record in the recordset
    Wend
    strTable = strTable & "</table>"
    FormatRecordsetToHTML = strTable
End Function