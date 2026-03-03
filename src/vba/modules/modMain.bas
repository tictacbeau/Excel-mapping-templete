Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
'  modMain — Application Entry Point & Orchestration
'
'  Handles:
'    - Workbook open/initialize
'    - Dashboard display
'    - Top-level processing workflow
'    - Menu/ribbon command routing
'    - Application-level error handling
' ============================================================

' ===========================================================
'  WORKBOOK OPEN — called from ThisWorkbook.Workbook_Open
' ===========================================================
Public Sub Application_Initialize()
    On Error GoTo InitErr

    ' Ensure system sheets exist
    modPayorRegistry.InitializeRegistry

    ' Initialize settings sheet
    modUtils.GetOrCreateSheet SHEET_SETTINGS, True
    modUtils.GetOrCreateSheet SHEET_AUDIT_LOG, True
    modUtils.GetOrCreateSheet SHEET_PROCESSING, True

    ' Set up dashboard sheet
    SetupDashboard

    ' Log startup
    modUtils.WriteAuditLog "STARTUP", "Application", APP_TITLE & " initialized.", , "", "OK"

    ' Show dashboard
    ShowDashboard

    Exit Sub
InitErr:
    MsgBox "Initialization error: " & Err.Description & vbCrLf & _
        "Please check that the workbook is not in read-only mode.", _
        vbCritical, APP_NAME
    Err.Clear
End Sub

' ===========================================================
'  DASHBOARD SETUP — configure the visible Dashboard sheet
' ===========================================================
Private Sub SetupDashboard()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_DASHBOARD)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SHEET_DASHBOARD
    End If

    ws.Visible = xlSheetVisible

    ' Clear and format
    ws.Cells.Clear
    ws.Cells.Font.Name = "Calibri"

    ' Title bar
    With ws.Range("A1:L1")
        .Merge
        .Value = APP_TITLE
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = UI_COLOR_PRIMARY
        .RowHeight = 36
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Subtitle
    With ws.Range("A2:L2")
        .Merge
        .Value = "Accounts Receivable Remittance Allocation Automation"
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = UI_COLOR_PRIMARY
        .RowHeight = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Section: Process Remittance
    AddDashboardSection ws, 4, 1, "PROCESS REMITTANCE", _
        "Process a new payment remittance file against a configured payor template.", _
        "Process_Remittance"

    ' Section: Manage Payors
    AddDashboardSection ws, 4, 5, "MANAGE PAYORS", _
        "Add, edit, delete, or duplicate payor configurations and templates.", _
        "Open_PayorManager"

    ' Section: Internal Data
    AddDashboardSection ws, 10, 1, "INTERNAL DATA", _
        "Import or refresh the internal invoice data (300K+ records from billing system).", _
        "Open_DataRefresh"

    ' Section: Audit Log
    AddDashboardSection ws, 10, 5, "AUDIT LOG", _
        "View the processing history, conversion log, and system events.", _
        "Open_AuditLog"

    ' Status bar at bottom
    With ws.Range("A16:L16")
        .Merge
        .Interior.Color = UI_COLOR_PANEL
        .Font.Size = 9
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
    UpdateStatusBar ws

    ' Column widths
    Dim c As Integer
    For c = 1 To 12
        ws.Columns(c).ColumnWidth = 10
    Next c
    ws.Columns(4).ColumnWidth = 2
    ws.Columns(8).ColumnWidth = 2

    ws.Activate
    ws.Cells(1, 1).Select
End Sub

Private Sub AddDashboardSection(ws As Worksheet, nRow As Integer, nCol As Integer, _
                                  sTitle As String, sDesc As String, sMacro As String)
    ' Section header
    With ws.Range(ws.Cells(nRow, nCol), ws.Cells(nRow, nCol + 2))
        .Merge
        .Value = sTitle
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = UI_COLOR_SECONDARY
        .RowHeight = 28
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Description
    With ws.Range(ws.Cells(nRow + 1, nCol), ws.Cells(nRow + 2, nCol + 2))
        .Merge
        .Value = sDesc
        .Font.Size = 9
        .Interior.Color = UI_COLOR_PANEL
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 40
    End With

    ' Action button (using a Shape/button-style cell)
    With ws.Range(ws.Cells(nRow + 3, nCol), ws.Cells(nRow + 3, nCol + 2))
        .Merge
        .Value = "Open >"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = UI_COLOR_PRIMARY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
End Sub

Private Sub UpdateStatusBar(ws As Worksheet)
    Dim nPayors As Long: nPayors = 0
    Dim nRecords As Long: nRecords = modDataLookup.GetDataModelRecordCount()
    On Error Resume Next
    Dim col As Collection
    Set col = modPayorRegistry.GetAllPayors()
    If Not col Is Nothing Then nPayors = col.Count
    On Error GoTo 0

    ws.Range("A16").Value = APP_NAME & " v" & APP_VERSION & _
        "   |   Payors: " & nPayors & _
        "   |   Internal Records: " & Format(nRecords, "#,##0") & _
        "   |   Last Updated: " & Format(Now(), "mm/dd/yyyy hh:mm")
End Sub

' ===========================================================
'  SHOW DASHBOARD — activate the dashboard sheet
' ===========================================================
Public Sub ShowDashboard()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_DASHBOARD)
    If Not ws Is Nothing Then
        ws.Activate
        UpdateStatusBar ws
    End If
    On Error GoTo 0
End Sub

' ===========================================================
'  PROCESS REMITTANCE — main processing entry point
'  Called from Dashboard button or ribbon command
' ===========================================================
Public Sub Process_Remittance()
    ' Check prerequisites
    Dim col As Collection
    Set col = modPayorRegistry.GetAllPayors()
    If col.Count = 0 Then
        modUtils.ShowInfo "No payors configured yet." & vbCrLf & _
            "Please set up at least one payor before processing.", APP_NAME
        Open_PayorManager
        Exit Sub
    End If

    ' Show the process remittance form
    frmProcessRemittance.Show
End Sub

' ===========================================================
'  OPEN PAYOR MANAGER
' ===========================================================
Public Sub Open_PayorManager()
    frmPayorManager.Show
End Sub

' ===========================================================
'  OPEN DATA REFRESH
' ===========================================================
Public Sub Open_DataRefresh()
    frmDataRefresh.Show
End Sub

' ===========================================================
'  OPEN AUDIT LOG — show the audit log sheet
' ===========================================================
Public Sub Open_AuditLog()
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_AUDIT_LOG, False)
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

' ===========================================================
'  OPEN PAYOR WIZARD — new payor setup
' ===========================================================
Public Sub Open_NewPayorWizard()
    Dim frm As New frmPayorWizard
    frm.StartNew
    frm.Show
End Sub

' ===========================================================
'  EDIT PAYOR — open wizard for existing payor
' ===========================================================
Public Sub Edit_Payor(sPayorId As String)
    Dim oPayor As clsPayor
    Set oPayor = modPayorRegistry.GetPayor(sPayorId)
    If oPayor Is Nothing Then
        modUtils.ShowError "Payor not found: " & sPayorId
        Exit Sub
    End If
    Dim frm As New frmPayorWizard
    frm.LoadPayor oPayor
    frm.Show
End Sub

' ===========================================================
'  DELETE PAYOR — with confirmation
' ===========================================================
Public Sub Delete_Payor(sPayorId As String, sPayorName As String)
    If Not modUtils.ConfirmAction( _
        "Are you sure you want to delete payor '" & sPayorName & "'?" & vbCrLf & _
        "This will also delete all associated templates. This cannot be undone.", _
        "Delete Payor") Then
        Exit Sub
    End If

    If modPayorRegistry.DeletePayor(sPayorId) Then
        modUtils.ShowInfo "Payor '" & sPayorName & "' deleted successfully."
    Else
        modUtils.ShowError "Failed to delete payor. Check the audit log for details."
    End If
End Sub

' ===========================================================
'  QUICK PROCESS — streamlined for repeat processing
'  Remembers last payor selection
' ===========================================================
Public Sub QuickProcess()
    Dim sLastPayor As String
    sLastPayor = modUtils.GetSetting(SETTING_LAST_PAYOR, "")

    If sLastPayor = "" Then
        Process_Remittance
        Exit Sub
    End If

    ' Open file picker directly
    Dim sFile As String
    sFile = Application.GetOpenFilename( _
        "Remittance Files (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", _
        1, "Select Remittance File for " & sLastPayor, , False)

    If sFile = "False" Or sFile = "" Then Exit Sub

    ' Find payor ID
    Dim oPayor As clsPayor
    Set oPayor = modPayorRegistry.GetPayorByName(sLastPayor)
    If oPayor Is Nothing Then
        Process_Remittance
        Exit Sub
    End If

    ' Process
    Application.StatusBar = "Processing remittance for " & sLastPayor & "..."
    Dim wsResult As Worksheet
    Set wsResult = modTemplateEngine.ProcessRemittance(oPayor.PayorId, sFile)
    Application.StatusBar = False

    If Not wsResult Is Nothing Then
        modUtils.ShowInfo "Processing complete!" & vbCrLf & _
            "Output sheet: " & wsResult.Name, "Success"
    End If
End Sub

' ===========================================================
'  EXPORT PAYOR CONFIG — save payor config as JSON file
' ===========================================================
Public Sub ExportPayorConfig(sPayorId As String)
    Dim oPayor As clsPayor
    Set oPayor = modPayorRegistry.GetPayor(sPayorId)
    If oPayor Is Nothing Then
        modUtils.ShowError "Payor not found: " & sPayorId
        Exit Sub
    End If

    Dim sSavePath As String
    sSavePath = Application.GetSaveAsFilename( _
        oPayor.PayorName & "_config.json", _
        "JSON Files (*.json),*.json", 1, "Export Payor Config")

    If sSavePath = "False" Then Exit Sub

    On Error Resume Next
    Dim fileNum As Integer: fileNum = FreeFile()
    Open sSavePath For Output As #fileNum
    Print #fileNum, oPayor.ToJson()
    Close #fileNum
    On Error GoTo 0

    modUtils.ShowInfo "Config exported to: " & sSavePath
End Sub

' ===========================================================
'  IMPORT PAYOR CONFIG — load from JSON file
' ===========================================================
Public Sub ImportPayorConfig()
    Dim sFile As String
    sFile = Application.GetOpenFilename( _
        "JSON Files (*.json),*.json", 1, "Import Payor Config", , False)
    If sFile = "False" Then Exit Sub

    Dim fileNum As Integer: fileNum = FreeFile()
    Dim sJson As String: sJson = ""
    On Error Resume Next
    Open sFile For Input As #fileNum
    Dim sLine As String
    Do While Not EOF(fileNum)
        Line Input #fileNum, sLine
        sJson = sJson & sLine
    Loop
    Close #fileNum
    On Error GoTo 0

    If sJson = "" Then
        modUtils.ShowError "Could not read config file: " & sFile
        Exit Sub
    End If

    Dim oPayor As New clsPayor
    oPayor.FromJson sJson
    oPayor.PayorId = ""  ' Force new ID

    Dim sErr As String: sErr = oPayor.Validate()
    If sErr <> "" Then
        modUtils.ShowError "Invalid config: " & sErr
        Exit Sub
    End If

    If modPayorRegistry.SavePayor(oPayor) Then
        modUtils.ShowInfo "Payor '" & oPayor.PayorName & "' imported successfully."
    Else
        modUtils.ShowError "Import failed. Check the audit log."
    End If
End Sub

' ===========================================================
'  ABOUT BOX
' ===========================================================
Public Sub ShowAbout()
    MsgBox APP_TITLE & vbCrLf & vbCrLf & _
        "Remittance Allocation Template Engine" & vbCrLf & _
        "Automates AR remittance breakdown processing." & vbCrLf & vbCrLf & _
        "Version: " & APP_VERSION & vbCrLf & _
        "Supports: XLSX, XLS (genuine + HTML), CSV input formats" & vbCrLf & _
        "Internal data: Excel Data Model (300K+ records)" & vbCrLf & vbCrLf & _
        "© " & Year(Now()) & " — All rights reserved.", _
        vbInformation, "About " & APP_NAME
End Sub
