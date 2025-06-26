' === OneClickBackupAndVersionControl ===
' Purpose:
'   - Create backups of the current workbook in chosen formats (.xlsm, .xlsx, .csv).
'   - Store backups in a "Backups" subfolder with timestamped filenames.
'   - Log each backup run in a visible "BackupLog" sheet for audit.

Option Explicit

Sub OneClickBackupAndVersionControl()
    ' --- Declare and initialize objects/variables ---
    Dim fso                As Object        ' FileSystemObject for folder checks
    Dim backupsFolderPath  As String        ' Full path to Backups folder
    Dim workbookFolderPath As String        ' Path where this workbook is saved
    Dim workbookName       As String        ' Name of this workbook without extension
    Dim currentTimestamp   As String        ' Timestamp string for versioning
    Dim backupFileBaseName As String        ' "workbook_backup_timestamp"
    Dim userPromptResult   As VbMsgBoxResult
    Dim chosenFormats      As String        ' Track which formats user selected

    ' --- Build paths and base file name ---
    workbookFolderPath = ThisWorkbook.Path & Application.PathSeparator
    workbookName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    currentTimestamp = Format(Now, "yyyymmdd_HHmmss")
    backupFileBaseName = workbookName & "_Backup_" & currentTimestamp
    backupsFolderPath = workbookFolderPath & "Backups" & Application.PathSeparator

    ' --- Ensure Backups folder exists ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(backupsFolderPath) Then
        fso.CreateFolder backupsFolderPath
    End If

    ' --- Prompt user for each format ---
    Dim saveAsXlsm  As Boolean
    Dim saveAsXlsx  As Boolean
    Dim saveAsCsv   As Boolean

    saveAsXlsm = (MsgBox("Create macro-enabled backup (.xlsm)?", vbYesNo + vbQuestion, "Backup Format") = vbYes)
    saveAsXlsx = (MsgBox("Create macro-free backup (.xlsx)?", vbYesNo + vbQuestion, "Backup Format") = vbYes)
    saveAsCsv = (MsgBox("Export active sheet as CSV (.csv)?", vbYesNo + vbQuestion, "Backup Format") = vbYes)

    Application.DisplayAlerts = False

    ' --- Backup .xlsm (full copy) ---
    If saveAsXlsm Then
        ThisWorkbook.SaveCopyAs backupsFolderPath & backupFileBaseName & ".xlsm"
        chosenFormats = chosenFormats & " .xlsm"
    End If

    ' --- Backup .xlsx (sheets only, no macros) ---
    If saveAsXlsx Then
        Dim tempWorkbook As Workbook
        ' Copy all sheets into new workbook (strips VBA)
        ThisWorkbook.Sheets.Copy
        Set tempWorkbook = ActiveWorkbook
        ' Save as .xlsx
        tempWorkbook.SaveAs fileName:=backupsFolderPath & backupFileBaseName & ".xlsx", FileFormat:=51
        tempWorkbook.Close SaveChanges:=False
        chosenFormats = chosenFormats & " .xlsx"
    End If

    ' --- Backup .csv (active sheet only) ---
    If saveAsCsv Then
        Dim csvWorkbook As Workbook
        ' Copy active sheet into new workbook
        ThisWorkbook.ActiveSheet.Copy
        Set csvWorkbook = ActiveWorkbook
        ' Save as UTF-8 CSV
        csvWorkbook.SaveAs fileName:=backupsFolderPath & backupFileBaseName & ".csv", FileFormat:=62
        csvWorkbook.Close SaveChanges:=False
        chosenFormats = chosenFormats & " .csv"
    End If

    Application.DisplayAlerts = True

    ' --- Log backup operation ---
    LogBackupRun currentTimestamp, backupFileBaseName, Trim(chosenFormats), Environ("Username"), backupsFolderPath

    ' --- Notify user ---
    MsgBox "Backup complete for formats:" & vbCrLf & chosenFormats & vbCrLf & _
           "Files saved in: " & backupsFolderPath, vbInformation, "Backup Success"
End Sub

' === LogBackupRun ===
' Parameters:
'    timestamp    - when backup ran
'    fileBaseName - base filename (without extension)
'    formats      - space-separated list of extensions saved
'    userName     - Windows user who ran backup
'    logFolder    - folder path where backups are stored
Sub LogBackupRun( _
    ByVal timestamp As String, _
    ByVal fileBaseName As String, _
    ByVal formats As String, _
    ByVal userName As String, _
    ByVal logFolder As String _
)
    Dim logWS   As Worksheet
    Dim logName As String: logName = "BackupLog"
    Dim nextRow As Long

    ' --- Locate or create the log sheet ---
    On Error Resume Next
    Set logWS = ThisWorkbook.Worksheets(logName)
    On Error GoTo 0

    If logWS Is Nothing Then
        ' Add sheet at end and make visible
        Set logWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        logWS.Name = logName
        logWS.Visible = xlSheetVisible
        ' Create header row
        With logWS
            .Range("A1:E1").Value = Array("Timestamp", "FileBaseName", "Formats", "User", "BackupPath")
            .Rows(1).Font.Bold = True
        End With
    Else
        ' Ensure it's visible for review
        logWS.Visible = xlSheetVisible
    End If

    ' --- Append new log entry ---
    nextRow = logWS.Cells(logWS.Rows.Count, "A").End(xlUp).Row + 1
    logWS.Cells(nextRow, 1).Value = timestamp
    logWS.Cells(nextRow, 2).Value = fileBaseName
    logWS.Cells(nextRow, 3).Value = formats
    logWS.Cells(nextRow, 4).Value = userName
    logWS.Cells(nextRow, 5).Value = logFolder
End Sub


