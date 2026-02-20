
Option Compare Database
Option Explicit

Public Function exportAllObjectsAsText()
    Dim stringDatabaseFilename As String
    
    Dim stringExportPath As String
    Dim fso As Object
    Set fso = Nothing
    Set fso = Nothing
    On Error GoTo exportAllObjectsAsText_Error
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim myComponent
    Dim sModuleType
    Dim sTempname
    
    
    Dim sOutstring
    stringDatabaseFilename = CurrentDb.Name
    Dim myType, myName, myPath, sStubADPFilename
    myType = fso.GetExtensionName(stringDatabaseFilename)
    myName = fso.GetBaseName(stringDatabaseFilename)
    myPath = fso.GetParentFolderName(stringDatabaseFilename)

    If (stringExportPath = "") Then
        stringExportPath = myPath & "\" & myName & "_ObjectsTextBackup\"
    End If

    SysCmd acSysCmdSetStatus, "Copy stub to " & sStubADPFilename & "..."
    On Error Resume Next
        fso.CreateFolder (stringExportPath)
    On Error GoTo 0
    SysCmd acSysCmdInitMeter, "Please wait while all forms are exported…", Application.CurrentProject.AllForms.Count
    Dim myObj As Object
    Dim integerCounter As Integer
    For Each myObj In Application.CurrentProject.AllForms
        integerCounter = integerCounter + 1
        Application.SaveAsText acForm, myObj.FullName, stringExportPath & "\" & myObj.FullName & ".form"
        Application.DoCmd.Close acForm, myObj.FullName
        SysCmd acSysCmdUpdateMeter, integerCounter
    Next
    SysCmd acSysCmdRemoveMeter: integerCounter = 0
    SysCmd acSysCmdInitMeter, "Please wait while all modules are exported", Application.CurrentProject.AllModules.Count
    For Each myObj In Application.CurrentProject.AllModules
        integerCounter = integerCounter + 1
        Application.SaveAsText acModule, myObj.FullName, stringExportPath & "\" & myObj.FullName & ".bas"
        SysCmd acSysCmdUpdateMeter, integerCounter
    Next
    SysCmd acSysCmdRemoveMeter: integerCounter = 0
    SysCmd acSysCmdInitMeter, "Please wait while all macros are exported", Application.CurrentProject.AllMacros.Count
    For Each myObj In Application.CurrentProject.AllMacros
        integerCounter = integerCounter + 1
        Application.SaveAsText acMacro, myObj.FullName, stringExportPath & "\" & myObj.FullName & ".macro"
        SysCmd acSysCmdUpdateMeter, integerCounter
    Next
    SysCmd acSysCmdRemoveMeter: integerCounter = 0
    SysCmd acSysCmdInitMeter, "Please wait while all reports are exported…", Application.CurrentProject.AllReports.Count
    For Each myObj In Application.CurrentProject.AllReports
        integerCounter = integerCounter + 1
        Application.SaveAsText acReport, myObj.FullName, stringExportPath & "\" & myObj.FullName & ".report"
        SysCmd acSysCmdUpdateMeter, integerCounter
    Next
    SysCmd acSysCmdRemoveMeter: integerCounter = 0
    Dim queryDefinition As DAO.QueryDef
    SysCmd acSysCmdInitMeter, "Please wait while this process runs…", CurrentDb.QueryDefs.Count
    For Each queryDefinition In CurrentDb.QueryDefs
        integerCounter = integerCounter + 1
        Dim stringSQLText As String
        Dim stringFilename As String
        stringFilename = stringExportPath & "\" & queryDefinition.Name & ".SQL"
        CreateTextFile stringFilename, queryDefinition.SQL
        SysCmd acSysCmdUpdateMeter, integerCounter
    Next
    SysCmd acSysCmdRemoveMeter
    LaunchExplorer stringExportPath
ExitHere:
   Exit Function

exportAllObjectsAsText_Error:
    Select Case Err.Number
    '1
    Case Else
        MsgBox "Unexpected Error:" _
            & vbCrLf & "Error " & Err.Number & " " & Err.Description _
            & " in procedure exportAllObjectsAsText of Module BackupDatabaseObjects" _
            , vbCritical, "Please Investigate"
        Resume ExitHere
        Resume 'For Debug Only
    End Select


End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.Source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function

Public Function CreateTextFile(stringFilename As String, stringContents As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(stringFilename)
    oFile.WriteLine stringContents
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Function


Public Function Test()
End Function