VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  clsTemplate — Template layout definition
'  Represents one template (SINGLE or MULTI) for one payor.
'  Contains an ordered collection of cell definitions (rows).
' ============================================================

' ---- Identity ----------------------------------------------
Public PayorId As String
Public TemplateType As String        ' "SINGLE" or "MULTI"
Public TemplateName As String        ' Display name

' ---- Section row counts ------------------------------------
Public HeaderRowCount As Integer     ' Number of header section rows
Public BodyRowCount As Integer       ' Number of body rows per invoice (template def)
Public FooterRowCount As Integer     ' Number of footer rows

' ---- Column definitions ------------------------------------
Public ColumnCount As Integer        ' Number of data columns
' Column widths array (parallel to column numbers)
Private m_ColWidths() As Single      ' -1 = no change, 0 = autofit, >0 = fixed
Private m_ColLabels() As String      ' Column header labels (for builder display)

' ---- Cell definitions collection ---------------------------
' Each entry is a serialized JSON string for one cell
Private m_Cells() As String          ' JSON cell definitions
Private m_CellCount As Integer

' ---- Body section first/last row in full template ----------
' Computed when building: first row index (1-based) of body section
Public BodyStartRow As Integer
Public BodyEndRow As Integer

' ---- Default formatting profiles ---------------------------
Public DefaultHeaderFormatJson As String
Public DefaultBodyFormatJson As String
Public DefaultFooterFormatJson As String
Public AlternateRowShading As Boolean  ' TRUE = alternate row fill in body

' ---- Print settings (JSON) ---------------------------------
Public PrintSettingsJson As String

' ---- Freeze pane settings ----------------------------------
Public FreezePaneRow As Integer        ' Row to freeze below (0 = no freeze)
Public FreezePaneCol As Integer        ' Column to freeze right of (0 = no freeze)

' ===========================================================
'  Initialize
' ===========================================================
Private Sub Class_Initialize()
    TemplateType = TMPL_SINGLE
    HeaderRowCount = 3
    BodyRowCount = 1
    FooterRowCount = 2
    ColumnCount = 8
    m_CellCount = 0
    ReDim m_Cells(0)
    ReDim m_ColWidths(0)
    ReDim m_ColLabels(0)
    AlternateRowShading = True
    FreezePaneRow = 3
    FreezePaneCol = 0

    ' Default formats (JSON)
    DefaultHeaderFormatJson = BuildDefaultHeaderJson()
    DefaultBodyFormatJson = BuildDefaultBodyJson()
    DefaultFooterFormatJson = BuildDefaultFooterJson()
    PrintSettingsJson = "{""orientation"":""portrait"",""fitToPage"":true,""fitToWidth"":1,""fitToHeight"":0}"
End Sub

' ===========================================================
'  Add a cell definition
' ===========================================================
Public Sub AddCell(sSection As String, nRowOffset As Integer, nCol As Integer, _
                    sCellType As String, sFieldMap As String, sDefaultValue As String, _
                    sFormatJson As String, bIsDynamic As Boolean, _
                    Optional sMergeRef As String = "", _
                    Optional sFormula As String = "")
    Dim j As String
    j = "{"
    j = j & """section"":""" & sSection & ""","
    j = j & """rowOffset"":" & CStr(nRowOffset) & ","
    j = j & """col"":" & CStr(nCol) & ","
    j = j & """cellType"":""" & sCellType & ""","
    j = j & """fieldMap"":""" & Esc(sFieldMap) & ""","
    j = j & """defaultValue"":""" & Esc(sDefaultValue) & ""","
    j = j & """formatJson"":" & IIf(sFormatJson = "", "{}", sFormatJson) & ","
    j = j & """isDynamic"":" & BoolJ(bIsDynamic) & ","
    j = j & """mergeRef"":""" & Esc(sMergeRef) & ""","
    j = j & """formula"":""" & Esc(sFormula) & """"
    j = j & "}"

    ReDim Preserve m_Cells(m_CellCount)
    m_Cells(m_CellCount) = j
    m_CellCount = m_CellCount + 1
End Sub

' ===========================================================
'  Cell count read-only
' ===========================================================
Public Property Get CellCount() As Integer
    CellCount = m_CellCount
End Property

Public Property Get CellJson(i As Integer) As String
    If i >= 0 And i < m_CellCount Then
        CellJson = m_Cells(i)
    Else
        CellJson = ""
    End If
End Property

' ===========================================================
'  Get cells filtered by section
' ===========================================================
Public Function GetSectionCells(sSection As String) As Collection
    Dim col As New Collection
    Dim i As Integer
    For i = 0 To m_CellCount - 1
        If modUtils.JsonGetString(m_Cells(i), "section", "") = sSection Then
            col.Add m_Cells(i)
        End If
    Next i
    Set GetSectionCells = col
End Function

' ===========================================================
'  Get body cells only (isDynamic=true cells)
' ===========================================================
Public Function GetBodyCells() As Collection
    Dim col As New Collection
    Dim i As Integer
    For i = 0 To m_CellCount - 1
        If modUtils.JsonGetBool(m_Cells(i), "isDynamic", False) Then
            col.Add m_Cells(i)
        End If
    Next i
    Set GetBodyCells = col
End Function

' ===========================================================
'  Column width helpers
' ===========================================================
Public Sub SetColumnWidth(nCol As Integer, dWidth As Single)
    If nCol > UBound(m_ColWidths) Then ReDim Preserve m_ColWidths(nCol)
    m_ColWidths(nCol - 1) = dWidth
End Sub

Public Function GetColumnWidth(nCol As Integer) As Single
    If nCol - 1 <= UBound(m_ColWidths) Then
        GetColumnWidth = m_ColWidths(nCol - 1)
    Else
        GetColumnWidth = -1
    End If
End Function

Public Sub SetColumnLabel(nCol As Integer, sLabel As String)
    If nCol > UBound(m_ColLabels) Then ReDim Preserve m_ColLabels(nCol)
    m_ColLabels(nCol - 1) = sLabel
End Sub

Public Function GetColumnLabel(nCol As Integer) As String
    If nCol - 1 <= UBound(m_ColLabels) Then
        GetColumnLabel = m_ColLabels(nCol - 1)
    Else
        GetColumnLabel = "Col" & CStr(nCol)
    End If
End Function

' ===========================================================
'  Total rows in the template for N invoices
' ===========================================================
Public Function TotalRows(nInvoices As Integer) As Integer
    Dim bodyRows As Integer
    If TemplateType = TMPL_MULTI Then
        bodyRows = BodyRowCount * nInvoices
    Else
        bodyRows = BodyRowCount
    End If
    TotalRows = HeaderRowCount + bodyRows + FooterRowCount
End Function

' ===========================================================
'  Serialize to JSON (stored in TemplateDB sheet)
' ===========================================================
Public Function ToJson() As String
    Dim j As String
    j = "{"
    j = j & """payorId"":""" & Esc(PayorId) & ""","
    j = j & """templateType"":""" & Esc(TemplateType) & ""","
    j = j & """templateName"":""" & Esc(TemplateName) & ""","
    j = j & """headerRowCount"":" & CStr(HeaderRowCount) & ","
    j = j & """bodyRowCount"":" & CStr(BodyRowCount) & ","
    j = j & """footerRowCount"":" & CStr(FooterRowCount) & ","
    j = j & """columnCount"":" & CStr(ColumnCount) & ","
    j = j & """bodyStartRow"":" & CStr(BodyStartRow) & ","
    j = j & """bodyEndRow"":" & CStr(BodyEndRow) & ","
    j = j & """alternateRowShading"":" & BoolJ(AlternateRowShading) & ","
    j = j & """freezePaneRow"":" & CStr(FreezePaneRow) & ","
    j = j & """freezePaneCol"":" & CStr(FreezePaneCol) & ","
    j = j & """defaultHeaderFormat"":" & DefaultHeaderFormatJson & ","
    j = j & """defaultBodyFormat"":" & DefaultBodyFormatJson & ","
    j = j & """defaultFooterFormat"":" & DefaultFooterFormatJson & ","
    j = j & """printSettings"":" & PrintSettingsJson & ","
    j = j & """cells"":" & CellsToJsonArray()
    j = j & "}"
    ToJson = j
End Function

' ===========================================================
'  Populate from JSON
' ===========================================================
Public Sub FromJson(sJson As String)
    PayorId = modUtils.JsonGetString(sJson, "payorId", "")
    TemplateType = modUtils.JsonGetString(sJson, "templateType", TMPL_SINGLE)
    TemplateName = modUtils.JsonGetString(sJson, "templateName", "")
    HeaderRowCount = CInt(modUtils.JsonGetNumber(sJson, "headerRowCount", 3))
    BodyRowCount = CInt(modUtils.JsonGetNumber(sJson, "bodyRowCount", 1))
    FooterRowCount = CInt(modUtils.JsonGetNumber(sJson, "footerRowCount", 2))
    ColumnCount = CInt(modUtils.JsonGetNumber(sJson, "columnCount", 8))
    BodyStartRow = CInt(modUtils.JsonGetNumber(sJson, "bodyStartRow", 0))
    BodyEndRow = CInt(modUtils.JsonGetNumber(sJson, "bodyEndRow", 0))
    AlternateRowShading = modUtils.JsonGetBool(sJson, "alternateRowShading", True)
    FreezePaneRow = CInt(modUtils.JsonGetNumber(sJson, "freezePaneRow", 3))
    FreezePaneCol = CInt(modUtils.JsonGetNumber(sJson, "freezePaneCol", 0))
    DefaultHeaderFormatJson = modUtils.JsonGetRawValue(sJson, "defaultHeaderFormat", BuildDefaultHeaderJson())
    DefaultBodyFormatJson = modUtils.JsonGetRawValue(sJson, "defaultBodyFormat", BuildDefaultBodyJson())
    DefaultFooterFormatJson = modUtils.JsonGetRawValue(sJson, "defaultFooterFormat", BuildDefaultFooterJson())
    PrintSettingsJson = modUtils.JsonGetRawValue(sJson, "printSettings", "{}")

    ' Parse cells array
    Dim sArr As String
    sArr = modUtils.JsonGetRawValue(sJson, "cells", "[]")
    ParseCellsFromJsonArray sArr
End Sub

' ===========================================================
'  Private helpers
' ===========================================================
Private Function CellsToJsonArray() As String
    Dim j As String: j = "["
    Dim i As Integer
    For i = 0 To m_CellCount - 1
        If i > 0 Then j = j & ","
        j = j & m_Cells(i)
    Next i
    j = j & "]"
    CellsToJsonArray = j
End Function

Private Sub ParseCellsFromJsonArray(sArr As String)
    m_CellCount = 0
    ReDim m_Cells(0)
    ' Extract JSON objects from array using simple bracket matching
    Dim items() As String
    items = modUtils.SplitJsonArray(sArr)
    Dim i As Integer
    For i = 0 To UBound(items)
        If Trim(items(i)) <> "" Then
            ReDim Preserve m_Cells(m_CellCount)
            m_Cells(m_CellCount) = Trim(items(i))
            m_CellCount = m_CellCount + 1
        End If
    Next i
End Sub

Private Function Esc(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    Esc = s
End Function

Private Function BoolJ(b As Boolean) As String
    BoolJ = IIf(b, "true", "false")
End Function

Private Function BuildDefaultHeaderJson() As String
    BuildDefaultHeaderJson = "{""fontName"":""Calibri"",""fontSize"":12,""fontColor"":16777215," & _
        """fontBold"":true,""fillColor"":3355187,""patternType"":1," & _
        """hAlign"":-4108,""vAlign"":-4108,""cellLocked"":true," & _
        """borderBottomStyle"":4,""borderBottomColor"":0}"
    ' fillColor 3355187 = #003366 (navy)
End Function

Private Function BuildDefaultBodyJson() As String
    BuildDefaultBodyJson = "{""fontName"":""Calibri"",""fontSize"":10,""fontColor"":0," & _
        """fontBold"":false,""fillColor"":-1,""patternType"":0," & _
        """hAlign"":2,""vAlign"":-4108,""cellLocked"":false," & _
        """borderBottomStyle"":1,""borderBottomColor"":12632256}"
End Function

Private Function BuildDefaultFooterJson() As String
    BuildDefaultFooterJson = "{""fontName"":""Calibri"",""fontSize"":10,""fontColor"":0," & _
        """fontBold"":true,""fillColor"":-1,""patternType"":0," & _
        """hAlign"":2,""vAlign"":-4108,""cellLocked"":true," & _
        """borderTopStyle"":4,""borderTopColor"":0}"
End Function
