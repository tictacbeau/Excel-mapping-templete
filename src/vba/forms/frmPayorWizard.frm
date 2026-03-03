VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPayorWizard
   Caption         =   "Payor Setup Wizard"
   ClientHeight    =   10800
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPayorWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  frmPayorWizard — Multi-step payor configuration wizard
'
'  Step 1: Payor Identity
'  Step 2: Remittance Format Detection
'  Step 3: Template Layout Builder
'  Step 4: Internal Data Linkage
'  Step 5: Review & Save
' ============================================================

Private m_CurrentStep As Integer
Private Const TOTAL_STEPS As Integer = 5
Private m_Payor As clsPayor
Private m_SingleTemplate As clsTemplate
Private m_MultiTemplate As clsTemplate
Private m_IsNew As Boolean

' ===========================================================
'  PUBLIC INTERFACE — called before Show
' ===========================================================

Public Sub StartNew()
    Set m_Payor = New clsPayor
    Set m_SingleTemplate = New clsTemplate
    Set m_MultiTemplate = New clsTemplate
    m_SingleTemplate.TemplateType = TMPL_SINGLE
    m_MultiTemplate.TemplateType = TMPL_MULTI
    m_IsNew = True
    m_CurrentStep = 1
End Sub

Public Sub LoadPayor(oPayor As clsPayor)
    Set m_Payor = oPayor
    Set m_SingleTemplate = modPayorRegistry.LoadTemplate(oPayor.PayorId, TMPL_SINGLE)
    Set m_MultiTemplate = modPayorRegistry.LoadTemplate(oPayor.PayorId, TMPL_MULTI)
    If m_SingleTemplate Is Nothing Then Set m_SingleTemplate = New clsTemplate
    If m_MultiTemplate Is Nothing Then Set m_MultiTemplate = New clsTemplate
    m_SingleTemplate.PayorId = oPayor.PayorId
    m_SingleTemplate.TemplateType = TMPL_SINGLE
    m_MultiTemplate.PayorId = oPayor.PayorId
    m_MultiTemplate.TemplateType = TMPL_MULTI
    m_IsNew = False
    m_CurrentStep = 1
End Sub

' ===========================================================
'  FORM INITIALIZE
' ===========================================================
Private Sub UserForm_Initialize()
    If m_Payor Is Nothing Then StartNew
    ShowStep m_CurrentStep
End Sub

' ===========================================================
'  STEP NAVIGATION
' ===========================================================
Private Sub btnNext_Click()
    If Not ValidateCurrentStep() Then Exit Sub
    SaveCurrentStep
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

Private Sub btnCancel_Click()
    If modUtils.ConfirmAction("Discard changes and close the wizard?") Then
        Me.Hide
    End If
End Sub

Private Sub btnSave_Click()
    ' Final step — save everything
    If Not ValidateCurrentStep() Then Exit Sub
    SaveCurrentStep

    ' Validate payor
    Dim sErr As String: sErr = m_Payor.Validate()
    If sErr <> "" Then
        MsgBox "Validation failed: " & sErr, vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Save payor
    If Not modPayorRegistry.SavePayor(m_Payor) Then
        MsgBox "Failed to save payor. Check audit log.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' Save templates
    m_SingleTemplate.PayorId = m_Payor.PayorId
    m_MultiTemplate.PayorId = m_Payor.PayorId
    modPayorRegistry.SaveTemplate m_SingleTemplate
    modPayorRegistry.SaveTemplate m_MultiTemplate

    MsgBox "Payor '" & m_Payor.PayorName & "' saved successfully!", _
           vbInformation, APP_NAME
    Me.Hide
End Sub

' ===========================================================
'  SHOW STEP — activate the appropriate panel
' ===========================================================
Private Sub ShowStep(nStep As Integer)
    ' Update step indicator
    lblStepIndicator.Caption = "Step " & nStep & " of " & TOTAL_STEPS

    ' Navigation buttons
    btnBack.Enabled = (nStep > 1)
    btnNext.Visible = (nStep < TOTAL_STEPS)
    btnSave.Visible = (nStep = TOTAL_STEPS)

    ' Hide all frames
    fraStep1.Visible = False
    fraStep2.Visible = False
    fraStep3.Visible = False
    fraStep4.Visible = False
    fraStep5.Visible = False

    ' Show active frame
    Select Case nStep
        Case 1: fraStep1.Visible = True: LoadStep1
        Case 2: fraStep2.Visible = True: LoadStep2
        Case 3: fraStep3.Visible = True: LoadStep3
        Case 4: fraStep4.Visible = True: LoadStep4
        Case 5: fraStep5.Visible = True: LoadStep5
    End Select
End Sub

' ===========================================================
'  STEP 1 — Payor Identity
' ===========================================================
Private Sub LoadStep1()
    txtPayorName.Value = m_Payor.PayorName
    ' Aliases
    Dim i As Integer
    Dim sAliases As String: sAliases = ""
    For i = 0 To m_Payor.AliasCount - 1
        If sAliases <> "" Then sAliases = sAliases & vbCrLf
        sAliases = sAliases & m_Payor.Aliases(i)
    Next i
    txtAliases.Value = sAliases

    ' File type dropdown
    cboFileType.Clear
    cboFileType.AddItem FILETYPE_XLSX & " (.xlsx)"
    cboFileType.AddItem FILETYPE_XLS_GENUINE & " (.xls - genuine binary)"
    cboFileType.AddItem FILETYPE_XLS_HTML & " (.xls - HTML table)"
    cboFileType.AddItem FILETYPE_CSV & " (.csv)"
    ' Set selection
    Dim ft As String: ft = m_Payor.ExpectedFileType
    Dim j As Integer
    For j = 0 To cboFileType.ListCount - 1
        If InStr(cboFileType.List(j), ft) > 0 Then
            cboFileType.ListIndex = j
            Exit For
        End If
    Next j

    txtNotes.Value = m_Payor.Notes
End Sub

Private Function ValidateStep1() As Boolean
    If Trim(txtPayorName.Value) = "" Then
        MsgBox "Payor name is required.", vbExclamation
        txtPayorName.SetFocus
        ValidateStep1 = False
        Exit Function
    End If
    ' Check for duplicate name (excluding current payor)
    If modPayorRegistry.PayorNameExists(Trim(txtPayorName.Value), m_Payor.PayorId) Then
        MsgBox "A payor named '" & Trim(txtPayorName.Value) & "' already exists.", vbExclamation
        txtPayorName.SetFocus
        ValidateStep1 = False
        Exit Function
    End If
    ValidateStep1 = True
End Function

Private Sub SaveStep1()
    m_Payor.PayorName = Trim(txtPayorName.Value)
    m_Payor.Notes = txtNotes.Value
    ' File type
    Dim sSelected As String: sSelected = cboFileType.Text
    If InStr(sSelected, FILETYPE_XLSX) > 0 Then m_Payor.ExpectedFileType = FILETYPE_XLSX
    If InStr(sSelected, FILETYPE_XLS_GENUINE) > 0 Then m_Payor.ExpectedFileType = FILETYPE_XLS_GENUINE
    If InStr(sSelected, FILETYPE_XLS_HTML) > 0 Then m_Payor.ExpectedFileType = FILETYPE_XLS_HTML
    If InStr(sSelected, FILETYPE_CSV) > 0 Then m_Payor.ExpectedFileType = FILETYPE_CSV
    ' Aliases
    m_Payor.AliasCount = 0
    ReDim m_Payor.Aliases(0)
    Dim lines() As String
    lines = Split(txtAliases.Value, vbCrLf)
    Dim i As Integer
    For i = 0 To UBound(lines)
        Dim s As String: s = Trim(lines(i))
        If s <> "" Then m_Payor.AddAlias s
    Next i
End Sub

' ===========================================================
'  STEP 2 — Remittance Format Detection
' ===========================================================
Private Sub LoadStep2()
    ' Invoice layout
    optLayoutSingleCell.Value = (m_Payor.InvoiceLayout = "SINGLE_CELL")
    optLayoutOnePerRow.Value = (m_Payor.InvoiceLayout = "ONE_PER_ROW")
    txtInvoiceDelimiter.Value = m_Payor.InvoiceDelimiter
    txtInvoiceRegex.Value = m_Payor.InvoiceRegex
    txtInvoiceColumn.Value = CStr(m_Payor.InvoiceColumn)
    txtAmountColumn.Value = CStr(m_Payor.AmountColumn)
    txtAmountDelimiter.Value = m_Payor.AmountDelimiter
    txtHeaderRowCount.Value = CStr(m_Payor.HeaderRowCount)
    txtInvoiceRow.Value = CStr(m_Payor.InvoiceRow)
    txtAmountRow.Value = CStr(m_Payor.AmountRow)
    cboDuplicateAction.Clear
    cboDuplicateAction.AddItem "FLAG - Mark duplicates with orange highlight"
    cboDuplicateAction.AddItem "SKIP - Ignore duplicate invoice numbers"
    cboDuplicateAction.AddItem "INCLUDE - Include all occurrences"
    Select Case m_Payor.DuplicateAction
        Case "FLAG": cboDuplicateAction.ListIndex = 0
        Case "SKIP": cboDuplicateAction.ListIndex = 1
        Case "INCLUDE": cboDuplicateAction.ListIndex = 2
    End Select
    UpdateLayoutControls
End Sub

Private Sub optLayoutSingleCell_Click()
    UpdateLayoutControls
End Sub

Private Sub optLayoutOnePerRow_Click()
    UpdateLayoutControls
End Sub

Private Sub UpdateLayoutControls()
    Dim bSingle As Boolean: bSingle = optLayoutSingleCell.Value
    ' Single-cell controls
    txtInvoiceDelimiter.Enabled = bSingle
    txtAmountDelimiter.Enabled = bSingle
    txtInvoiceRow.Enabled = bSingle
    txtAmountRow.Enabled = bSingle
    ' One-per-row controls
    txtInvoiceColumn.Enabled = Not bSingle
    txtAmountColumn.Enabled = Not bSingle
End Sub

Private Sub btnTestRegex_Click()
    Dim sPattern As String: sPattern = Trim(txtInvoiceRegex.Value)
    Dim sSample As String: sSample = Trim(txtRegexSample.Value)

    If sPattern = "" Then
        MsgBox "Enter a regex pattern first.", vbInformation
        Exit Sub
    End If
    If sSample = "" Then
        MsgBox "Enter sample text to test against.", vbInformation
        Exit Sub
    End If

    If Not modUtils.IsValidRegex(sPattern) Then
        lblRegexResult.Caption = "ERROR: Invalid regex pattern!"
        lblRegexResult.ForeColor = RGB(200, 0, 0)
        Exit Sub
    End If

    Dim sResult As String: sResult = modUtils.RegexTest(sPattern, sSample)
    lblRegexResult.Caption = "Matches: " & sResult
    If sResult = "NO MATCH" Then
        lblRegexResult.ForeColor = RGB(200, 0, 0)
    Else
        lblRegexResult.ForeColor = RGB(0, 128, 0)
    End If
End Sub

Private Function ValidateStep2() As Boolean
    If optLayoutSingleCell.Value Then
        If Trim(txtInvoiceDelimiter.Value) = "" And Trim(txtInvoiceRegex.Value) = "" Then
            MsgBox "Single-cell layout requires either a delimiter or regex pattern.", vbExclamation
            ValidateStep2 = False: Exit Function
        End If
    Else
        If Val(txtInvoiceColumn.Value) < 1 Then
            MsgBox "Invoice column must be at least 1.", vbExclamation
            ValidateStep2 = False: Exit Function
        End If
    End If
    If Trim(txtInvoiceRegex.Value) <> "" And Not modUtils.IsValidRegex(txtInvoiceRegex.Value) Then
        MsgBox "The regex pattern is not valid: " & txtInvoiceRegex.Value, vbExclamation
        ValidateStep2 = False: Exit Function
    End If
    ValidateStep2 = True
End Function

Private Sub SaveStep2()
    m_Payor.InvoiceLayout = IIf(optLayoutSingleCell.Value, "SINGLE_CELL", "ONE_PER_ROW")
    m_Payor.InvoiceDelimiter = txtInvoiceDelimiter.Value
    m_Payor.InvoiceRegex = Trim(txtInvoiceRegex.Value)
    m_Payor.InvoiceColumn = CInt(Val(txtInvoiceColumn.Value))
    m_Payor.AmountColumn = CInt(Val(txtAmountColumn.Value))
    m_Payor.AmountDelimiter = txtAmountDelimiter.Value
    m_Payor.HeaderRowCount = CInt(Val(txtHeaderRowCount.Value))
    m_Payor.InvoiceRow = CInt(Val(txtInvoiceRow.Value))
    m_Payor.AmountRow = CInt(Val(txtAmountRow.Value))
    Select Case cboDuplicateAction.ListIndex
        Case 0: m_Payor.DuplicateAction = "FLAG"
        Case 1: m_Payor.DuplicateAction = "SKIP"
        Case 2: m_Payor.DuplicateAction = "INCLUDE"
    End Select
End Sub

' ===========================================================
'  STEP 3 — Template Layout Builder (launches separate form)
' ===========================================================
Private Sub LoadStep3()
    lblTemplateSummary.Caption = GetTemplateSummary()
End Sub

Private Sub btnEditSingleTemplate_Click()
    Dim frm As New frmTemplateBuilder
    frm.LoadTemplate m_SingleTemplate, m_Payor
    frm.Show
    Set m_SingleTemplate = frm.GetTemplate()
    lblTemplateSummary.Caption = GetTemplateSummary()
End Sub

Private Sub btnEditMultiTemplate_Click()
    Dim frm As New frmTemplateBuilder
    frm.LoadTemplate m_MultiTemplate, m_Payor
    frm.Show
    Set m_MultiTemplate = frm.GetTemplate()
    lblTemplateSummary.Caption = GetTemplateSummary()
End Sub

Private Function GetTemplateSummary() As String
    Dim s As String
    s = "Single-Invoice Template: "
    s = s & IIf(m_SingleTemplate.CellCount > 0, _
        m_SingleTemplate.CellCount & " cell(s) defined", "Not configured") & vbCrLf
    s = s & "Multi-Invoice Template: "
    s = s & IIf(m_MultiTemplate.CellCount > 0, _
        m_MultiTemplate.CellCount & " cell(s) defined", "Not configured")
    GetTemplateSummary = s
End Function

Private Function ValidateStep3() As Boolean
    ' Templates are optional at this step (can be configured later)
    ValidateStep3 = True
End Function

Private Sub SaveStep3()
    ' Templates saved separately via their builder form
End Sub

' ===========================================================
'  STEP 4 — Internal Data Linkage
' ===========================================================
Private Sub LoadStep4()
    ' Load current lookup config
    Dim oConfig As Object
    Set oConfig = modUtils.SimpleJsonToDictionary(m_Payor.LookupConfigJson)

    If oConfig.Exists("primaryKey") Then
        txtPrimaryKey.Value = oConfig("primaryKey")
    Else
        txtPrimaryKey.Value = "InvoiceNumber"
    End If

    If oConfig.Exists("matchType") Then
        cboMatchType.Text = oConfig("matchType")
    Else
        cboMatchType.ListIndex = 0  ' exact
    End If

    If oConfig.Exists("noMatchAction") Then
        cboNoMatchAction.Text = oConfig("noMatchAction")
    End If

    ' Populate column list from Data Model
    Dim cols() As String: cols = modDataLookup.GetDataModelColumns()
    lstDataColumns.Clear
    Dim i As Integer
    For i = 0 To UBound(cols)
        If cols(i) <> "" Then lstDataColumns.AddItem cols(i)
    Next i

    ' Data model status
    Dim nRec As Long: nRec = modDataLookup.GetDataModelRecordCount()
    lblDataModelStatus.Caption = "Internal data: " & Format(nRec, "#,##0") & " records loaded."
    If nRec = 0 Then
        lblDataModelStatus.ForeColor = RGB(200, 0, 0)
    Else
        lblDataModelStatus.ForeColor = RGB(0, 128, 0)
    End If
End Sub

Private Sub cboMatchType_Change()
    ' Update UI based on match type
End Sub

Private Function ValidateStep4() As Boolean
    If Trim(txtPrimaryKey.Value) = "" Then
        MsgBox "Primary key field is required for data lookup.", vbExclamation
        ValidateStep4 = False: Exit Function
    End If
    ValidateStep4 = True
End Function

Private Sub SaveStep4()
    Dim sMatchType As String
    sMatchType = IIf(cboMatchType.ListIndex >= 0, cboMatchType.Text, "exact")
    Dim sNoMatch As String
    sNoMatch = IIf(cboNoMatchAction.ListIndex >= 0, cboNoMatchAction.Text, "flag")

    ' Build selected columns list
    Dim sReturnCols As String: sReturnCols = "["
    Dim i As Integer
    For i = 0 To lstDataColumns.ListCount - 1
        If lstDataColumns.Selected(i) Then
            If sReturnCols <> "[" Then sReturnCols = sReturnCols & ","
            sReturnCols = sReturnCols & """" & lstDataColumns.List(i) & """"
        End If
    Next i
    sReturnCols = sReturnCols & "]"

    m_Payor.LookupConfigJson = "{" & _
        """primaryKey"":""" & Trim(txtPrimaryKey.Value) & """," & _
        """matchType"":""" & sMatchType & """," & _
        """noMatchAction"":""" & sNoMatch & """," & _
        """returnColumns"":" & sReturnCols & _
        "}"
End Sub

' ===========================================================
'  STEP 5 — Review & Save
' ===========================================================
Private Sub LoadStep5()
    ' Build review summary
    Dim s As String
    s = "PAYOR CONFIGURATION REVIEW" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
    s = s & "Name: " & m_Payor.PayorName & vbCrLf
    s = s & "File Type: " & m_Payor.ExpectedFileType & vbCrLf
    s = s & "Invoice Layout: " & m_Payor.InvoiceLayout & vbCrLf
    If m_Payor.InvoiceRegex <> "" Then
        s = s & "Regex: " & m_Payor.InvoiceRegex & vbCrLf
    End If
    s = s & "Duplicate Action: " & m_Payor.DuplicateAction & vbCrLf
    s = s & vbCrLf
    s = s & "Single Template: " & m_SingleTemplate.CellCount & " cell(s)" & vbCrLf
    s = s & "Multi Template: " & m_MultiTemplate.CellCount & " cell(s)" & vbCrLf
    s = s & vbCrLf
    If m_Payor.LookupConfigJson <> "{}" Then
        s = s & "Lookup Key: " & modUtils.JsonGetString(m_Payor.LookupConfigJson, "primaryKey", "") & vbCrLf
    End If
    If m_Payor.Notes <> "" Then
        s = s & vbCrLf & "Notes: " & m_Payor.Notes & vbCrLf
    End If
    txtReview.Value = s
End Sub

Private Function ValidateStep5() As Boolean
    ValidateStep5 = True
End Function

Private Sub SaveStep5()
    ' Nothing to save — btnSave handles the actual persistence
End Sub

' ===========================================================
'  VALIDATION ROUTING
' ===========================================================
Private Function ValidateCurrentStep() As Boolean
    Select Case m_CurrentStep
        Case 1: ValidateCurrentStep = ValidateStep1()
        Case 2: ValidateCurrentStep = ValidateStep2()
        Case 3: ValidateCurrentStep = ValidateStep3()
        Case 4: ValidateCurrentStep = ValidateStep4()
        Case 5: ValidateCurrentStep = ValidateStep5()
    End Select
End Function

Private Sub SaveCurrentStep()
    Select Case m_CurrentStep
        Case 1: SaveStep1
        Case 2: SaveStep2
        Case 3: SaveStep3
        Case 4: SaveStep4
        Case 5: SaveStep5
    End Select
End Sub

' ---- Controls summary: ----
' lblStepIndicator    — "Step X of 5"
' fraStep1..fraStep5  — Step frames (one visible at a time)
' btnBack, btnNext, btnSave, btnCancel
'
' Step 1: txtPayorName, txtAliases, cboFileType, txtNotes
' Step 2: optLayoutSingleCell, optLayoutOnePerRow,
'         txtInvoiceDelimiter, txtAmountDelimiter,
'         txtInvoiceRegex, txtRegexSample, btnTestRegex, lblRegexResult,
'         txtInvoiceColumn, txtAmountColumn,
'         txtHeaderRowCount, txtInvoiceRow, txtAmountRow,
'         cboDuplicateAction
' Step 3: lblTemplateSummary, btnEditSingleTemplate, btnEditMultiTemplate
' Step 4: txtPrimaryKey, cboMatchType, cboNoMatchAction,
'         lstDataColumns, lblDataModelStatus
' Step 5: txtReview
