VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataRefresh
   Caption         =   "Internal Data Refresh"
   ClientHeight    =   8500
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   9000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataRefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  frmDataRefresh — Internal Data Model Refresh Wizard
'
'  Allows the user to:
'    1. Select a new source file (xlsx/csv)
'    2. Preview schema / column names
'    3. Run the refresh (replaces existing data)
'    4. View before/after statistics
'    5. Rollback if needed
' ============================================================

Private m_SourceFile As String
Private m_CurrentStep As Integer
Private Const TOTAL_STEPS As Integer = 3

Private Sub UserForm_Initialize()
    m_CurrentStep = 1
    LoadCurrentStatus
    ShowStep m_CurrentStep
End Sub

Private Sub LoadCurrentStatus()
    Dim nRec As Long: nRec = modDataLookup.GetDataModelRecordCount()
    Dim sLastSource As String
    sLastSource = modUtils.GetSetting(SETTING_DATA_SOURCE_PATH, "")

    lblCurrentStatus.Caption = "Current Data Model:" & vbCrLf & _
        "  Records: " & Format(nRec, "#,##0") & vbCrLf & _
        "  Last source: " & IIf(sLastSource <> "", sLastSource, "(none)") & vbCrLf & _
        "  Status: " & IIf(modDataLookup.IsDataModelLoaded, "Loaded", "Empty")

    If nRec = 0 Then
        lblCurrentStatus.ForeColor = RGB(200, 100, 0)
    Else
        lblCurrentStatus.ForeColor = RGB(0, 128, 0)
    End If

    ' Schema info
    txtSchemaInfo.Value = modDataLookup.ValidateDataModelSchema()
End Sub

' ===========================================================
'  STEP NAVIGATION
' ===========================================================
Private Sub ShowStep(nStep As Integer)
    lblStepIndicator.Caption = "Step " & nStep & " of " & TOTAL_STEPS
    btnNext.Visible = (nStep < TOTAL_STEPS)
    btnBack.Enabled = (nStep > 1)
    btnRunRefresh.Visible = (nStep = TOTAL_STEPS)

    fraStep1.Visible = (nStep = 1)
    fraStep2.Visible = (nStep = 2)
    fraStep3.Visible = (nStep = 3)

    Select Case nStep
        Case 2: PreviewSourceFile
        Case 3: PrepareRefreshConfirm
    End Select
End Sub

Private Sub btnNext_Click()
    If Not ValidateCurrentStep() Then Exit Sub
    If m_CurrentStep < TOTAL_STEPS Then
        m_CurrentStep = m_CurrentStep + 1
        ShowStep m_CurrentStep
    End If
End Sub

Private Sub btnBack_Click()
    If m_CurrentStep > 1 Then
        m_CurrentStep = m_CurrentStep - 1
        ShowStep m_CurrentStep
    End If
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

' ===========================================================
'  STEP 1 — Select Source File
' ===========================================================
Private Sub btnBrowse_Click()
    Dim sFile As String
    sFile = Application.GetOpenFilename( _
        "Data Files (*.xlsx;*.xlsm;*.csv),*.xlsx;*.xlsm;*.csv", _
        1, "Select Internal Data Source File", , False)

    If sFile <> "False" And sFile <> "" Then
        m_SourceFile = sFile
        txtSourceFile.Value = sFile
        lblBrowseStatus.Caption = "File: " & Right(sFile, 60)
        lblBrowseStatus.ForeColor = RGB(0, 128, 0)
    End If
End Sub

Private Sub btnUseLastSource_Click()
    Dim sLast As String
    sLast = modUtils.GetSetting(SETTING_DATA_SOURCE_PATH, "")
    If sLast <> "" And modUtils.FileExists(sLast) Then
        m_SourceFile = sLast
        txtSourceFile.Value = sLast
        lblBrowseStatus.Caption = "Using last source: " & Right(sLast, 60)
        lblBrowseStatus.ForeColor = RGB(0, 128, 0)
    Else
        lblBrowseStatus.Caption = "No previous source file found."
        lblBrowseStatus.ForeColor = RGB(200, 0, 0)
    End If
End Sub

Private Function ValidateStep1() As Boolean
    If m_SourceFile = "" Then
        MsgBox "Please select a source file.", vbExclamation
        ValidateStep1 = False: Exit Function
    End If
    If Not modUtils.FileExists(m_SourceFile) Then
        MsgBox "File not found: " & m_SourceFile, vbExclamation
        ValidateStep1 = False: Exit Function
    End If
    ValidateStep1 = True
End Function

' ===========================================================
'  STEP 2 — Preview Source Schema
' ===========================================================
Private Sub PreviewSourceFile()
    lblPreviewStatus.Caption = "Loading preview..."
    Me.Repaint

    On Error GoTo PreviewErr
    ' Open source to peek at columns
    Application.DisplayAlerts = False
    Dim wbPreview As Workbook
    Set wbPreview = Application.Workbooks.Open(m_SourceFile, ReadOnly:=True)
    Dim ws As Worksheet: Set ws = wbPreview.Sheets(1)

    Dim nCols As Integer: nCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim nRows As Long: nRows = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim sPreview As String
    sPreview = "Source File Preview:" & vbCrLf
    sPreview = sPreview & "  Columns: " & nCols & vbCrLf
    sPreview = sPreview & "  Rows (including header): " & nRows & vbCrLf
    sPreview = sPreview & "  Data rows: " & (nRows - 1) & vbCrLf & vbCrLf
    sPreview = sPreview & "Column Names:" & vbCrLf

    Dim c As Integer
    For c = 1 To nCols
        sPreview = sPreview & "  " & c & ". " & Trim(ws.Cells(1, c).Value) & vbCrLf
    Next c

    ' Show first few data rows
    sPreview = sPreview & vbCrLf & "Sample Data (first 3 rows):" & vbCrLf
    Dim r As Long
    For r = 2 To IIf(nRows >= 4, 4, nRows)
        Dim sRow As String: sRow = "  Row " & (r - 1) & ": "
        For c = 1 To IIf(nCols > 5, 5, nCols)
            sRow = sRow & "[" & Left(CStr(ws.Cells(r, c).Value), 15) & "] "
        Next c
        sPreview = sPreview & sRow & vbCrLf
    Next r

    wbPreview.Close False
    Application.DisplayAlerts = True
    txtPreview.Value = sPreview
    lblPreviewStatus.Caption = "Preview loaded successfully."
    lblPreviewStatus.ForeColor = RGB(0, 128, 0)
    Exit Sub

PreviewErr:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbPreview Is Nothing Then wbPreview.Close False
    On Error GoTo 0
    txtPreview.Value = "Error reading file: " & Err.Description
    lblPreviewStatus.Caption = "Preview failed."
    lblPreviewStatus.ForeColor = RGB(200, 0, 0)
    Err.Clear
End Sub

Private Function ValidateStep2() As Boolean
    ValidateStep2 = True  ' Preview is optional
End Function

' ===========================================================
'  STEP 3 — Confirm and Run Refresh
' ===========================================================
Private Sub PrepareRefreshConfirm()
    Dim nCurrent As Long: nCurrent = modDataLookup.GetDataModelRecordCount()
    txtConfirmInfo.Value = _
        "You are about to refresh the internal data model." & vbCrLf & vbCrLf & _
        "Source file: " & m_SourceFile & vbCrLf & _
        "Current records: " & Format(nCurrent, "#,##0") & vbCrLf & vbCrLf & _
        "IMPORTANT:" & vbCrLf & _
        "• The current data will be REPLACED entirely." & vbCrLf & _
        "• A backup will be made automatically (rollback available)." & vbCrLf & _
        "• All lookup caches will be cleared." & vbCrLf & _
        "• Processing formulas will recalculate on next use." & vbCrLf & vbCrLf & _
        "Click 'Run Refresh' to proceed."
End Sub

Private Function ValidateStep3() As Boolean
    ValidateStep3 = True
End Function

Private Sub btnRunRefresh_Click()
    btnRunRefresh.Enabled = False
    lblRefreshStatus.Caption = "Refreshing data... please wait."
    lblRefreshStatus.ForeColor = RGB(0, 0, 200)
    Me.Repaint
    Application.StatusBar = "Importing internal data from: " & Right(m_SourceFile, 40) & "..."

    Dim nBefore As Long: nBefore = modDataLookup.GetDataModelRecordCount()
    Dim bSuccess As Boolean
    bSuccess = modDataLookup.RefreshDataModel(m_SourceFile)
    Application.StatusBar = False

    Dim nAfter As Long: nAfter = modDataLookup.GetDataModelRecordCount()

    If bSuccess Then
        lblRefreshStatus.Caption = "Refresh complete!"
        lblRefreshStatus.ForeColor = RGB(0, 128, 0)
        txtRefreshResult.Value = _
            "Refresh Results:" & vbCrLf & _
            "  Records before: " & Format(nBefore, "#,##0") & vbCrLf & _
            "  Records after: " & Format(nAfter, "#,##0") & vbCrLf & _
            "  Change: " & IIf(nAfter >= nBefore, "+", "") & (nAfter - nBefore) & vbCrLf & vbCrLf & _
            "A backup of the previous data has been saved." & vbCrLf & _
            "Click 'Rollback' if you need to revert."
        btnRollback.Enabled = True
        MsgBox "Data refresh complete!" & vbCrLf & _
            "Before: " & Format(nBefore, "#,##0") & " records" & vbCrLf & _
            "After: " & Format(nAfter, "#,##0") & " records", _
            vbInformation, APP_NAME
    Else
        lblRefreshStatus.Caption = "Refresh failed! See audit log."
        lblRefreshStatus.ForeColor = RGB(200, 0, 0)
        txtRefreshResult.Value = "Refresh failed. Previous data is intact. Check audit log."
    End If

    btnRunRefresh.Enabled = True
    LoadCurrentStatus
End Sub

Private Sub btnRollback_Click()
    If Not modUtils.ConfirmAction( _
        "Roll back to the previous data backup?" & vbCrLf & _
        "This will replace current data with the last backup.") Then Exit Sub

    If modDataLookup.RollbackDataModel() Then
        MsgBox "Rollback successful! Previous data restored.", vbInformation, APP_NAME
        LoadCurrentStatus
        lblRefreshStatus.Caption = "Rollback complete."
        lblRefreshStatus.ForeColor = RGB(0, 128, 0)
    Else
        MsgBox "Rollback failed. No backup available.", vbCritical, APP_NAME
    End If
End Sub

Private Sub btnPQRefresh_Click()
    ' Refresh the Power Query connection if configured
    modDataLookup.RefreshPowerQueryConnection
    LoadCurrentStatus
    MsgBox "Power Query connection refreshed.", vbInformation, APP_NAME
End Sub

' ---- Controls summary: ----
' lblStepIndicator    — "Step X of 3"
' lblCurrentStatus    — Current data model status
' txtSchemaInfo       — Current schema column list
' fraStep1..fraStep3  — Step frames
'
' Step 1: txtSourceFile, btnBrowse, btnUseLastSource, lblBrowseStatus
' Step 2: txtPreview, lblPreviewStatus
' Step 3: txtConfirmInfo, btnRunRefresh, lblRefreshStatus,
'         txtRefreshResult, btnRollback, btnPQRefresh
'
' btnNext, btnBack, btnClose
