Attribute VB_Name = "modThisWorkbook"
Option Explicit

' ============================================================
'  ThisWorkbook Module Stubs
'
'  These event handlers belong in the ThisWorkbook class module.
'  Copy these handlers into the ThisWorkbook module after
'  importing VBA code into your .xlsm file.
'
'  DO NOT import this as a regular .bas module.
'  This file documents the events needed for proper init.
' ============================================================

' --- Copy into ThisWorkbook module ---

' Private Sub Workbook_Open()
'     modMain.Application_Initialize
' End Sub
'
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     ' Save workbook if needed
'     On Error Resume Next
'     If Not ThisWorkbook.Saved Then
'         Dim answer As VbMsgBoxResult
'         answer = MsgBox("Save changes to " & APP_NAME & " workbook?", _
'                         vbYesNoCancel + vbQuestion, APP_NAME)
'         If answer = vbYes Then
'             ThisWorkbook.Save
'         ElseIf answer = vbCancel Then
'             Cancel = True
'         End If
'     End If
'     On Error GoTo 0
' End Sub
'
' Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'     ' Route dashboard button clicks to macro handlers
'     Dim sSheet As String: sSheet = Sh.Name
'     If sSheet = SHEET_DASHBOARD Then
'         HandleDashboardClick Sh, Target
'     End If
' End Sub
'
' Private Sub HandleDashboardClick(ws As Worksheet, rng As Range)
'     ' Check if clicked on an action cell
'     Dim sVal As String: sVal = Trim(rng.Value)
'     If sVal = "Open >" Then
'         Dim r As Long: r = rng.Row
'         If r = 7 Then modMain.Process_Remittance
'         If r = 7 And rng.Column >= 5 Then modMain.Open_PayorManager
'         If r = 13 Then modMain.Open_DataRefresh
'         If r = 13 And rng.Column >= 5 Then modMain.Open_AuditLog
'     End If
' End Sub
