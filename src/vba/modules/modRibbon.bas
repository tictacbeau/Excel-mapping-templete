Attribute VB_Name = "modRibbon"
Option Explicit

' ============================================================
'  modRibbon — Custom Ribbon Callbacks
'  All ribbon button onAction handlers route to modMain
' ============================================================

' Ribbon object (needed for Invalidate)
Private m_Ribbon As IRibbonUI

' Called when ribbon loads
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    On Error GoTo RibbonLoadErr
    Set m_Ribbon = ribbon
    Exit Sub
RibbonLoadErr:
    ' Non-fatal: ribbon object unavailable
End Sub

' Invalidate ribbon (force re-evaluate getEnabled/getVisible)
Public Sub InvalidateRibbon()
    If Not m_Ribbon Is Nothing Then m_Ribbon.Invalidate
End Sub

' ============================================================
'  RIBBON BUTTON HANDLERS
' ============================================================

Public Sub Ribbon_ProcessRemittance(control As IRibbonControl)
    modMain.Process_Remittance
End Sub

Public Sub Ribbon_QuickProcess(control As IRibbonControl)
    modMain.QuickProcess
End Sub

Public Sub Ribbon_ShowDashboard(control As IRibbonControl)
    modMain.ShowDashboard
End Sub

Public Sub Ribbon_ManagePayors(control As IRibbonControl)
    modMain.Open_PayorManager
End Sub

Public Sub Ribbon_NewPayor(control As IRibbonControl)
    modMain.Open_NewPayorWizard
End Sub

Public Sub Ribbon_ImportPayorConfig(control As IRibbonControl)
    modMain.ImportPayorConfig
End Sub

Public Sub Ribbon_RefreshData(control As IRibbonControl)
    modMain.Open_DataRefresh
End Sub

Public Sub Ribbon_DataModelStatus(control As IRibbonControl)
    Dim sStatus As String
    sStatus = modDataLookup.ValidateDataModelSchema()
    sStatus = sStatus & vbCrLf & "Cached records: " & _
        Format(modDataLookup.GetDataModelRecordCount(), "#,##0")
    modUtils.ShowInfo sStatus, "Internal Data Model Status"
End Sub

Public Sub Ribbon_InvalidateCache(control As IRibbonControl)
    modDataLookup.InvalidateCache
    modUtils.ShowInfo "Lookup cache cleared. Data will reload on next processing run.", _
        "Cache Cleared"
End Sub

Public Sub Ribbon_AuditLog(control As IRibbonControl)
    modMain.Open_AuditLog
End Sub

Public Sub Ribbon_RegexTester(control As IRibbonControl)
    ' Simple regex tester — uses InputBox chain
    Dim sPattern As String
    sPattern = InputBox( _
        "Enter a regex pattern to test:" & vbCrLf & _
        "Examples: \d{7}  |  INV-\d{6}  |  [A-Z]{2}\d{5}", _
        "Regex Pattern Tester", "")
    If sPattern = "" Then Exit Sub

    If Not modUtils.IsValidRegex(sPattern) Then
        modUtils.ShowError "Invalid regex pattern: " & sPattern
        Exit Sub
    End If

    Dim sSample As String
    sSample = InputBox( _
        "Enter sample text to test against:" & vbCrLf & _
        "Example: 1234567, 1234568, 1234569", _
        "Sample Text", "")
    If sSample = "" Then Exit Sub

    Dim sResult As String
    sResult = modUtils.RegexTest(sPattern, sSample)
    MsgBox "Pattern: " & sPattern & vbCrLf & vbCrLf & _
           "Matches: " & sResult, vbInformation, "Regex Test Result"
End Sub

Public Sub Ribbon_About(control As IRibbonControl)
    modMain.ShowAbout
End Sub
