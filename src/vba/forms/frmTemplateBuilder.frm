VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplateBuilder
   Caption         =   "Template Layout Builder"
   ClientHeight    =   12000
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   16000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplateBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  frmTemplateBuilder — Interactive Template Layout Builder
'
'  Provides:
'    - Grid-based cell mapper (header/body/footer sections)
'    - Cell type assignment (DATA/LOOKUP/CALC/BLANK/LABEL)
'    - Field mapping dropdown
'    - Formatting panel (font, fill, borders, number format)
'    - Live preview with color-coded cell types
'    - Section row count configuration
' ============================================================

Private m_Template As clsTemplate
Private m_Payor As clsPayor
Private m_SelectedRow As Integer
Private m_SelectedCol As Integer
Private m_GridCells() As String         ' JSON format per grid position
Private m_MaxRows As Integer
Private m_MaxCols As Integer

Private Const MAX_GRID_ROWS As Integer = 30
Private Const MAX_GRID_COLS As Integer = 12

' ===========================================================
'  PUBLIC INTERFACE
' ===========================================================
Public Sub LoadTemplate(oTemplate As clsTemplate, oPayor As clsPayor)
    Set m_Template = oTemplate
    Set m_Payor = oPayor
End Sub

Public Function GetTemplate() As clsTemplate
    Set GetTemplate = m_Template
End Function

' ===========================================================
'  FORM INITIALIZE
' ===========================================================
Private Sub UserForm_Initialize()
    If m_Template Is Nothing Then Set m_Template = New clsTemplate
    If m_Payor Is Nothing Then Set m_Payor = New clsPayor

    m_MaxRows = MAX_GRID_ROWS
    m_MaxCols = MAX_GRID_COLS
    ReDim m_GridCells(m_MaxRows, m_MaxCols)

    LoadTemplateToGrid
    SetupCellTypeDropdown
    SetupFieldMapDropdown
    SetupFormatControls
    lblTitle.Caption = "Template Builder — " & _
        m_Payor.PayorName & " (" & m_Template.TemplateType & ")"
    RefreshGridPreview
    SelectCell 1, 1
End Sub

' ===========================================================
'  LOAD TEMPLATE DATA INTO GRID
' ===========================================================
Private Sub LoadTemplateToGrid()
    ' Initialize all cells as BLANK
    Dim r As Integer, c As Integer
    For r = 1 To m_MaxRows
        For c = 1 To m_MaxCols
            m_GridCells(r, c) = BuildBlankCellJson(r, c)
        Next c
    Next r

    ' Overlay template cell definitions
    Dim i As Integer
    For i = 0 To m_Template.CellCount - 1
        Dim sCell As String: sCell = m_Template.CellJson(i)
        Dim sSection As String: sSection = modUtils.JsonGetString(sCell, "section", SECTION_HEADER)
        Dim nOffset As Integer: nOffset = CInt(modUtils.JsonGetNumber(sCell, "rowOffset", 1))
        Dim nCol As Integer: nCol = CInt(modUtils.JsonGetNumber(sCell, "col", 1))

        Dim nGridRow As Integer
        Select Case sSection
            Case SECTION_HEADER: nGridRow = nOffset
            Case SECTION_BODY: nGridRow = m_Template.HeaderRowCount + nOffset
            Case SECTION_FOOTER: nGridRow = m_Template.HeaderRowCount + m_Template.BodyRowCount + nOffset
        End Select

        If nGridRow >= 1 And nGridRow <= m_MaxRows And nCol >= 1 And nCol <= m_MaxCols Then
            m_GridCells(nGridRow, nCol) = sCell
        End If
    Next i

    ' Set section row controls
    txtHeaderRows.Value = CStr(m_Template.HeaderRowCount)
    txtBodyRows.Value = CStr(m_Template.BodyRowCount)
    txtFooterRows.Value = CStr(m_Template.FooterRowCount)
    txtColumnCount.Value = CStr(m_Template.ColumnCount)
    chkAlternateShading.Value = m_Template.AlternateRowShading
    txtFreezeRow.Value = CStr(m_Template.FreezePaneRow)
End Sub

Private Function BuildBlankCellJson(r As Integer, c As Integer) As String
    BuildBlankCellJson = "{""section"":""" & GetSectionForRow(r) & """," & _
        """rowOffset"":" & GetRowOffset(r) & "," & _
        """col"":" & c & "," & _
        """cellType"":""" & CELLTYPE_BLANK & """," & _
        """fieldMap"":"""",""defaultValue"":"""",""formatJson"":{}," & _
        """isDynamic"":" & IIf(GetSectionForRow(r) = SECTION_BODY, "true", "false") & "," & _
        """mergeRef"":"""",""formula"":""""}"
End Function

Private Function GetSectionForRow(r As Integer) As String
    Dim nH As Integer: nH = CInt(Val(txtHeaderRows.Value))
    Dim nB As Integer: nB = CInt(Val(txtBodyRows.Value))
    If r <= nH Then
        GetSectionForRow = SECTION_HEADER
    ElseIf r <= nH + nB Then
        GetSectionForRow = SECTION_BODY
    Else
        GetSectionForRow = SECTION_FOOTER
    End If
End Function

Private Function GetRowOffset(r As Integer) As Integer
    Dim nH As Integer: nH = CInt(Val(txtHeaderRows.Value))
    Dim nB As Integer: nB = CInt(Val(txtBodyRows.Value))
    If r <= nH Then
        GetRowOffset = r
    ElseIf r <= nH + nB Then
        GetRowOffset = r - nH
    Else
        GetRowOffset = r - nH - nB
    End If
End Function

' ===========================================================
'  CELL SELECTION
' ===========================================================
Private Sub SelectCell(nRow As Integer, nCol As Integer)
    m_SelectedRow = nRow
    m_SelectedCol = nCol

    Dim sCell As String: sCell = m_GridCells(nRow, nCol)
    lblSelectedCell.Caption = modUtils.CellAddress(nRow, nCol) & " (" & GetSectionForRow(nRow) & ")"

    ' Load cell type
    cboCellType.Text = modUtils.JsonGetString(sCell, "cellType", CELLTYPE_BLANK)

    ' Load field map
    txtFieldMap.Value = modUtils.JsonGetString(sCell, "fieldMap", "")

    ' Load default value
    txtDefaultValue.Value = modUtils.JsonGetString(sCell, "defaultValue", "")

    ' Load formula
    txtFormula.Value = modUtils.JsonGetString(sCell, "formula", "")

    ' Load formatting
    Dim sFormat As String: sFormat = modUtils.JsonGetRawValue(sCell, "formatJson", "{}")
    If sFormat <> "{}" And sFormat <> "" Then
        LoadFormatPanel sFormat
    Else
        LoadFormatPanelDefaults GetSectionForRow(nRow)
    End If

    UpdateCellTypeControls
    RefreshGridPreview
End Sub

' ===========================================================
'  CELL TYPE CHANGED
' ===========================================================
Private Sub cboCellType_Change()
    UpdateCellTypeControls
    ApplyCellChanges
End Sub

Private Sub UpdateCellTypeControls()
    Dim sCT As String: sCT = cboCellType.Text
    ' Show/hide controls based on cell type
    lblFieldMap.Visible = (sCT = CELLTYPE_DATA Or sCT = CELLTYPE_LOOKUP)
    cboFieldMap.Visible = (sCT = CELLTYPE_DATA Or sCT = CELLTYPE_LOOKUP)
    lblFormula.Visible = (sCT = CELLTYPE_CALC)
    txtFormula.Visible = (sCT = CELLTYPE_CALC)
    lblDefaultValue.Visible = (sCT = CELLTYPE_LABEL Or sCT = CELLTYPE_CALC)
    txtDefaultValue.Visible = (sCT = CELLTYPE_LABEL Or sCT = CELLTYPE_CALC)
End Sub

' ===========================================================
'  APPLY CHANGES to current selected cell
' ===========================================================
Private Sub ApplyCellChanges()
    If m_SelectedRow < 1 Or m_SelectedCol < 1 Then Exit Sub

    Dim sSection As String: sSection = GetSectionForRow(m_SelectedRow)
    Dim nOffset As Integer: nOffset = GetRowOffset(m_SelectedRow)
    Dim bDynamic As Boolean: bDynamic = (sSection = SECTION_BODY)
    Dim sCT As String: sCT = cboCellType.Text
    Dim sField As String: sField = cboFieldMap.Text
    Dim sDefault As String: sDefault = txtDefaultValue.Value
    Dim sFormula As String: sFormula = txtFormula.Value
    Dim sFormatJson As String: sFormatJson = GetFormatJson()

    m_GridCells(m_SelectedRow, m_SelectedCol) = _
        "{""section"":""" & sSection & """," & _
        """rowOffset"":" & nOffset & "," & _
        """col"":" & m_SelectedCol & "," & _
        """cellType"":""" & sCT & """," & _
        """fieldMap"":""" & sField & """," & _
        """defaultValue"":""" & Replace(sDefault, """", "\""") & """," & _
        """formatJson"":" & sFormatJson & "," & _
        """isDynamic"":" & IIf(bDynamic, "true", "false") & "," & _
        """mergeRef"":""" & txtMergeRef.Value & """," & _
        """formula"":""" & Replace(sFormula, """", "\""") & """}"

    RefreshGridPreview
End Sub

' ===========================================================
'  SAVE TEMPLATE from grid
' ===========================================================
Private Sub btnSave_Click()
    ' Update section counts from UI
    m_Template.HeaderRowCount = CInt(Val(txtHeaderRows.Value))
    m_Template.BodyRowCount = CInt(Val(txtBodyRows.Value))
    m_Template.FooterRowCount = CInt(Val(txtFooterRows.Value))
    m_Template.ColumnCount = CInt(Val(txtColumnCount.Value))
    m_Template.AlternateRowShading = chkAlternateShading.Value
    m_Template.FreezePaneRow = CInt(Val(txtFreezeRow.Value))

    ' Clear existing cells and rebuild from grid
    ' We rebuild clsTemplate from m_GridCells
    Dim newTemplate As New clsTemplate
    newTemplate.PayorId = m_Template.PayorId
    newTemplate.TemplateType = m_Template.TemplateType
    newTemplate.HeaderRowCount = m_Template.HeaderRowCount
    newTemplate.BodyRowCount = m_Template.BodyRowCount
    newTemplate.FooterRowCount = m_Template.FooterRowCount
    newTemplate.ColumnCount = m_Template.ColumnCount
    newTemplate.AlternateRowShading = m_Template.AlternateRowShading
    newTemplate.FreezePaneRow = m_Template.FreezePaneRow

    Dim r As Integer, c As Integer
    Dim nH As Integer: nH = m_Template.HeaderRowCount
    Dim nB As Integer: nB = m_Template.BodyRowCount
    Dim nTotalRows As Integer: nTotalRows = nH + nB + m_Template.FooterRowCount

    For r = 1 To nTotalRows
        For c = 1 To m_Template.ColumnCount
            Dim sCell As String: sCell = m_GridCells(r, c)
            Dim sCT As String: sCT = modUtils.JsonGetString(sCell, "cellType", CELLTYPE_BLANK)
            ' Skip truly empty/blank cells unless explicitly marked
            ' Include all non-blank cells
            If sCT <> CELLTYPE_BLANK Or modUtils.JsonGetString(sCell, "fieldMap", "") <> "" Then
                newTemplate.AddCell _
                    modUtils.JsonGetString(sCell, "section", GetSectionForRow(r)), _
                    CInt(modUtils.JsonGetNumber(sCell, "rowOffset", GetRowOffset(r))), _
                    c, _
                    sCT, _
                    modUtils.JsonGetString(sCell, "fieldMap", ""), _
                    modUtils.JsonGetString(sCell, "defaultValue", ""), _
                    modUtils.JsonGetRawValue(sCell, "formatJson", "{}"), _
                    modUtils.JsonGetBool(sCell, "isDynamic", GetSectionForRow(r) = SECTION_BODY), _
                    modUtils.JsonGetString(sCell, "mergeRef", ""), _
                    modUtils.JsonGetString(sCell, "formula", "")
            End If
        Next c
    Next r

    Set m_Template = newTemplate
    MsgBox "Template saved (" & m_Template.CellCount & " cell definitions).", vbInformation, APP_NAME
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    If modUtils.ConfirmAction("Discard template changes?") Then
        Me.Hide
    End If
End Sub

' ===========================================================
'  REFRESH GRID PREVIEW — color-code cells by type
' ===========================================================
Private Sub RefreshGridPreview()
    ' The actual grid is rendered on a hidden processing sheet for preview
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_PROCESSING, True)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    Dim nH As Integer: nH = CInt(Val(txtHeaderRows.Value))
    Dim nB As Integer: nB = CInt(Val(txtBodyRows.Value))
    Dim nF As Integer: nF = CInt(Val(txtFooterRows.Value))
    Dim nC As Integer: nC = CInt(Val(txtColumnCount.Value))
    Dim nTotal As Integer: nTotal = nH + nB + nF

    Dim r As Integer, c As Integer
    For r = 1 To nTotal
        For c = 1 To nC
            Dim sCell As String: sCell = m_GridCells(r, c)
            Dim sCT As String: sCT = modUtils.JsonGetString(sCell, "cellType", CELLTYPE_BLANK)
            Dim oCell As Range: Set oCell = ws.Cells(r, c)

            ' Color by cell type
            Select Case sCT
                Case CELLTYPE_DATA: oCell.Interior.Color = COLOR_DATA
                Case CELLTYPE_LOOKUP: oCell.Interior.Color = COLOR_LOOKUP
                Case CELLTYPE_CALC: oCell.Interior.Color = COLOR_CALC
                Case CELLTYPE_BLANK: oCell.Interior.Color = COLOR_BLANK
                Case CELLTYPE_LABEL: oCell.Interior.Color = COLOR_LABEL
            End Select

            ' Show field map or label text as cell value
            Dim sDisplay As String
            Select Case sCT
                Case CELLTYPE_LABEL: sDisplay = modUtils.JsonGetString(sCell, "defaultValue", "")
                Case CELLTYPE_DATA: sDisplay = "[" & modUtils.JsonGetString(sCell, "fieldMap", "DATA") & "]"
                Case CELLTYPE_LOOKUP: sDisplay = "{" & modUtils.JsonGetString(sCell, "fieldMap", "LOOKUP") & "}"
                Case CELLTYPE_CALC: sDisplay = "=" & Left(modUtils.JsonGetString(sCell, "formula", "CALC"), 15)
                Case Else: sDisplay = ""
            End Select
            oCell.Value = sDisplay
            oCell.Font.Size = 8

            ' Section separators
            If r = nH + 1 Then  ' First body row
                oCell.Borders(xlEdgeTop).LineStyle = xlMedium
                oCell.Borders(xlEdgeTop).Color = RGB(0, 0, 100)
            End If
            If r = nH + nB + 1 Then  ' First footer row
                oCell.Borders(xlEdgeTop).LineStyle = xlMedium
                oCell.Borders(xlEdgeTop).Color = RGB(0, 0, 100)
            End If

            ' Highlight selected cell
            If r = m_SelectedRow And c = m_SelectedCol Then
                oCell.Borders.LineStyle = xlMedium
                oCell.Borders.Color = RGB(255, 0, 0)
            End If
        Next c
    Next r

    ' Section labels in column nC + 1
    ws.Cells(CInt(nH / 2) + 1, nC + 2).Value = "HEADER"
    ws.Cells(nH + CInt(nB / 2) + 1, nC + 2).Value = "BODY"
    ws.Cells(nH + nB + CInt(nF / 2) + 1, nC + 2).Value = "FOOTER"
    ws.Columns.AutoFit
End Sub

' ===========================================================
'  FORMATTING PANEL
' ===========================================================
Private Sub SetupFormatControls()
    ' Font family
    cboFontName.Clear
    cboFontName.AddItem "Calibri"
    cboFontName.AddItem "Arial"
    cboFontName.AddItem "Times New Roman"
    cboFontName.AddItem "Courier New"
    cboFontName.Text = "Calibri"

    ' Number format
    cboNumberFormat.Clear
    cboNumberFormat.AddItem "General"
    cboNumberFormat.AddItem "$#,##0.00"
    cboNumberFormat.AddItem "#,##0"
    cboNumberFormat.AddItem "0.00%"
    cboNumberFormat.AddItem "MM/DD/YYYY"
    cboNumberFormat.AddItem "@"
    cboNumberFormat.Text = "General"

    ' H-Align
    cboHAlign.Clear
    cboHAlign.AddItem "Left"
    cboHAlign.AddItem "Center"
    cboHAlign.AddItem "Right"
    cboHAlign.ListIndex = 0

    ' V-Align
    cboVAlign.Clear
    cboVAlign.AddItem "Top"
    cboVAlign.AddItem "Middle"
    cboVAlign.AddItem "Bottom"
    cboVAlign.ListIndex = 1

    txtFontSize.Value = "10"
    txtFontColor.Value = "#000000"
    txtFillColor.Value = ""
    chkBold.Value = False
    chkItalic.Value = False
    chkUnderline.Value = False
    chkWrapText.Value = False
    chkCellLocked.Value = False
End Sub

Private Sub LoadFormatPanel(sFormatJson As String)
    Dim oFmt As New clsCellFormat
    oFmt.FromJson sFormatJson
    cboFontName.Text = oFmt.FontName
    txtFontSize.Value = CStr(oFmt.FontSize)
    txtFontColor.Value = modUtils.LongToHex(oFmt.FontColor)
    txtFillColor.Value = IIf(oFmt.FillColor = -1, "", modUtils.LongToHex(oFmt.FillColor))
    chkBold.Value = oFmt.FontBold
    chkItalic.Value = oFmt.FontItalic
    chkUnderline.Value = oFmt.FontUnderline
    chkWrapText.Value = oFmt.WrapText
    chkCellLocked.Value = oFmt.CellLocked
    cboNumberFormat.Text = oFmt.NumberFormat
    Select Case oFmt.HAlign
        Case xlLeft: cboHAlign.ListIndex = 0
        Case xlCenter: cboHAlign.ListIndex = 1
        Case xlRight: cboHAlign.ListIndex = 2
    End Select
    Select Case oFmt.VAlign
        Case xlTop: cboVAlign.ListIndex = 0
        Case xlCenter: cboVAlign.ListIndex = 1
        Case xlBottom: cboVAlign.ListIndex = 2
    End Select
End Sub

Private Sub LoadFormatPanelDefaults(sSection As String)
    Dim oFmt As New clsCellFormat
    Select Case sSection
        Case SECTION_HEADER: oFmt.SetHeaderDefaults
        Case SECTION_BODY: ' defaults from Class_Initialize
        Case SECTION_FOOTER:
            oFmt.FontBold = True
            oFmt.BorderTopStyle = xlMedium
    End Select
    LoadFormatPanel oFmt.ToJson()
End Sub

Private Function GetFormatJson() As String
    Dim oFmt As New clsCellFormat
    oFmt.FontName = cboFontName.Text
    oFmt.FontSize = CSng(Val(txtFontSize.Value))
    If txtFontColor.Value <> "" Then oFmt.FontColor = modUtils.HexToLong(txtFontColor.Value)
    oFmt.FontBold = chkBold.Value
    oFmt.FontItalic = chkItalic.Value
    oFmt.FontUnderline = chkUnderline.Value
    oFmt.WrapText = chkWrapText.Value
    oFmt.CellLocked = chkCellLocked.Value
    oFmt.NumberFormat = cboNumberFormat.Text
    Select Case cboHAlign.ListIndex
        Case 0: oFmt.HAlign = xlLeft
        Case 1: oFmt.HAlign = xlCenter
        Case 2: oFmt.HAlign = xlRight
    End Select
    Select Case cboVAlign.ListIndex
        Case 0: oFmt.VAlign = xlTop
        Case 1: oFmt.VAlign = xlCenter
        Case 2: oFmt.VAlign = xlBottom
    End Select
    If txtFillColor.Value <> "" Then
        oFmt.FillColor = modUtils.HexToLong(txtFillColor.Value)
        oFmt.PatternType = xlSolid
    End If
    oFmt.MergeRange = txtMergeRef.Value
    GetFormatJson = oFmt.ToJson()
End Function

' ===========================================================
'  FIELD MAP DROPDOWN
' ===========================================================
Private Sub SetupFieldMapDropdown()
    cboFieldMap.Clear
    ' Standard remittance fields (DATA type)
    cboFieldMap.AddItem "invoice_number"
    cboFieldMap.AddItem "invoice_amount"
    cboFieldMap.AddItem "invoice_date"
    cboFieldMap.AddItem "payment_date"
    cboFieldMap.AddItem "payment_amount"
    cboFieldMap.AddItem "check_number"
    cboFieldMap.AddItem "payor_name"
    cboFieldMap.AddItem "remittance_id"
    cboFieldMap.AddItem "total_extracted"
    cboFieldMap.AddItem "variance"
    cboFieldMap.AddItem "invoice_count"
    cboFieldMap.AddItem "row_number"

    ' Add Data Model columns (LOOKUP type)
    Dim cols() As String
    cols = modDataLookup.GetDataModelColumns()
    Dim i As Integer
    For i = 0 To UBound(cols)
        If cols(i) <> "" Then cboFieldMap.AddItem cols(i)
    Next i
End Sub

Private Sub SetupCellTypeDropdown()
    cboCellType.Clear
    cboCellType.AddItem CELLTYPE_LABEL & " — Static text / header label"
    cboCellType.AddItem CELLTYPE_DATA & " — Remittance data (blue)"
    cboCellType.AddItem CELLTYPE_LOOKUP & " — Internal DB lookup (green)"
    cboCellType.AddItem CELLTYPE_CALC & " — Formula / calculated (yellow)"
    cboCellType.AddItem CELLTYPE_BLANK & " — Intentionally blank (white)"
    cboCellType.ListIndex = 0
End Sub

' Apply format button
Private Sub btnApplyFormat_Click()
    ApplyCellChanges
End Sub

' Preset: Header format
Private Sub btnPresetHeader_Click()
    LoadFormatPanelDefaults SECTION_HEADER
    ApplyCellChanges
End Sub

' Preset: Currency
Private Sub btnPresetCurrency_Click()
    cboNumberFormat.Text = "$#,##0.00"
    cboHAlign.ListIndex = 2  ' Right
    ApplyCellChanges
End Sub

' Preset: Footer totals
Private Sub btnPresetFooter_Click()
    LoadFormatPanelDefaults SECTION_FOOTER
    ApplyCellChanges
End Sub

' ---- Controls summary: ----
' lblTitle               — "Template Builder — PayorName (SINGLE/MULTI)"
' lblSelectedCell        — "A1 (HEADER)" etc.
' cboCellType            — Cell type dropdown
' cboFieldMap/txtFieldMap — Field mapping
' txtDefaultValue        — Label text or default value
' txtFormula             — Formula string
' txtMergeRef            — Merge range reference
' txtHeaderRows, txtBodyRows, txtFooterRows, txtColumnCount
' chkAlternateShading    — Alternate row shading
' txtFreezeRow           — Freeze pane row
' cboFontName, txtFontSize, txtFontColor, txtFillColor
' chkBold, chkItalic, chkUnderline, chkWrapText, chkCellLocked
' cboNumberFormat, cboHAlign, cboVAlign
' btnApplyFormat, btnPresetHeader, btnPresetCurrency, btnPresetFooter
' btnSave, btnCancel
' lblFieldMap, lblFormula, lblDefaultValue
