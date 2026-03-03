VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  clsCellFormat — Cell formatting properties
'  Encapsulates all Format Cells dialog properties
' ============================================================

' ---- Font Properties ---------------------------------------
Public FontName As String          ' e.g. "Calibri"
Public FontSize As Single          ' e.g. 10
Public FontColor As Long           ' Excel Long color
Public FontBold As Boolean
Public FontItalic As Boolean
Public FontUnderline As Boolean
Public FontStrikethrough As Boolean

' ---- Fill / Background ------------------------------------
Public FillColor As Long           ' Excel Long color; -1 = no fill
Public PatternType As Long         ' xlNone, xlSolid, etc.

' ---- Number Format ----------------------------------------
Public NumberFormat As String      ' e.g. "$#,##0.00", "MM/DD/YYYY", "General"

' ---- Alignment --------------------------------------------
' Horizontal: xlLeft=2, xlCenter=-4108, xlRight=-4152
Public HAlign As Long
' Vertical: xlTop=-4160, xlCenter=-4108, xlBottom=-4107
Public VAlign As Long
Public WrapText As Boolean
Public Indent As Integer           ' Indent level (0-15)
Public Orientation As Integer      ' Text rotation in degrees (-90 to 90)

' ---- Borders ----------------------------------------------
' Border sides: xlEdgeTop=8, xlEdgeBottom=9, xlEdgeLeft=7, xlEdgeRight=10
' Style: xlContinuous=1, xlDash=2, xlDot=4, xlDouble=4, xlMedium=-4138, xlThick=4
Public BorderTopStyle As Long
Public BorderTopColor As Long
Public BorderBottomStyle As Long
Public BorderBottomColor As Long
Public BorderLeftStyle As Long
Public BorderLeftColor As Long
Public BorderRightStyle As Long
Public BorderRightColor As Long

' ---- Merge ------------------------------------------------
Public MergeRange As String        ' e.g. "A1:D1"; empty = no merge

' ---- Column / Row sizing (applied at column/row level) ---
Public ColumnWidth As Single       ' 0 = autofit; -1 = no change
Public RowHeight As Single         ' 0 = autofit; -1 = no change

' ---- Protection -------------------------------------------
Public CellLocked As Boolean       ' TRUE = locked when sheet protected

' ---- Conditional Formatting Rules (serialized JSON array) -
Public ConditionalFormats As String  ' JSON e.g. [{"condition":"<0","fillColor":"#FF0000"}]

' ===========================================================
'  Initialize with defaults
' ===========================================================
Private Sub Class_Initialize()
    FontName = "Calibri"
    FontSize = 10
    FontColor = 0             ' Black
    FontBold = False
    FontItalic = False
    FontUnderline = False
    FontStrikethrough = False

    FillColor = -1            ' No fill
    PatternType = xlNone

    NumberFormat = "General"

    HAlign = xlLeft
    VAlign = xlCenter
    WrapText = False
    Indent = 0
    Orientation = 0

    BorderTopStyle = xlNone
    BorderTopColor = 0
    BorderBottomStyle = xlNone
    BorderBottomColor = 0
    BorderLeftStyle = xlNone
    BorderLeftColor = 0
    BorderRightStyle = xlNone
    BorderRightColor = 0

    MergeRange = ""
    ColumnWidth = -1
    RowHeight = -1
    CellLocked = False
    ConditionalFormats = "[]"
End Sub

' ===========================================================
'  Serialize to JSON string
' ===========================================================
Public Function ToJson() As String
    Dim j As String
    j = "{"
    j = j & """fontName"":""" & Escape(FontName) & ""","
    j = j & """fontSize"":" & CStr(FontSize) & ","
    j = j & """fontColor"":" & CStr(FontColor) & ","
    j = j & """fontBold"":" & BoolToJson(FontBold) & ","
    j = j & """fontItalic"":" & BoolToJson(FontItalic) & ","
    j = j & """fontUnderline"":" & BoolToJson(FontUnderline) & ","
    j = j & """fontStrikethrough"":" & BoolToJson(FontStrikethrough) & ","
    j = j & """fillColor"":" & CStr(FillColor) & ","
    j = j & """patternType"":" & CStr(PatternType) & ","
    j = j & """numberFormat"":""" & Escape(NumberFormat) & ""","
    j = j & """hAlign"":" & CStr(HAlign) & ","
    j = j & """vAlign"":" & CStr(VAlign) & ","
    j = j & """wrapText"":" & BoolToJson(WrapText) & ","
    j = j & """indent"":" & CStr(Indent) & ","
    j = j & """orientation"":" & CStr(Orientation) & ","
    j = j & """borderTopStyle"":" & CStr(BorderTopStyle) & ","
    j = j & """borderTopColor"":" & CStr(BorderTopColor) & ","
    j = j & """borderBottomStyle"":" & CStr(BorderBottomStyle) & ","
    j = j & """borderBottomColor"":" & CStr(BorderBottomColor) & ","
    j = j & """borderLeftStyle"":" & CStr(BorderLeftStyle) & ","
    j = j & """borderLeftColor"":" & CStr(BorderLeftColor) & ","
    j = j & """borderRightStyle"":" & CStr(BorderRightStyle) & ","
    j = j & """borderRightColor"":" & CStr(BorderRightColor) & ","
    j = j & """mergeRange"":""" & Escape(MergeRange) & ""","
    j = j & """columnWidth"":" & CStr(ColumnWidth) & ","
    j = j & """rowHeight"":" & CStr(RowHeight) & ","
    j = j & """cellLocked"":" & BoolToJson(CellLocked) & ","
    j = j & """conditionalFormats"":" & ConditionalFormats
    j = j & "}"
    ToJson = j
End Function

' ===========================================================
'  Populate from JSON string (simple key-value parser)
' ===========================================================
Public Sub FromJson(sJson As String)
    FontName = modUtils.JsonGetString(sJson, "fontName", "Calibri")
    FontSize = CSng(modUtils.JsonGetNumber(sJson, "fontSize", 10))
    FontColor = CLng(modUtils.JsonGetNumber(sJson, "fontColor", 0))
    FontBold = modUtils.JsonGetBool(sJson, "fontBold", False)
    FontItalic = modUtils.JsonGetBool(sJson, "fontItalic", False)
    FontUnderline = modUtils.JsonGetBool(sJson, "fontUnderline", False)
    FontStrikethrough = modUtils.JsonGetBool(sJson, "fontStrikethrough", False)
    FillColor = CLng(modUtils.JsonGetNumber(sJson, "fillColor", -1))
    PatternType = CLng(modUtils.JsonGetNumber(sJson, "patternType", xlNone))
    NumberFormat = modUtils.JsonGetString(sJson, "numberFormat", "General")
    HAlign = CLng(modUtils.JsonGetNumber(sJson, "hAlign", xlLeft))
    VAlign = CLng(modUtils.JsonGetNumber(sJson, "vAlign", xlCenter))
    WrapText = modUtils.JsonGetBool(sJson, "wrapText", False)
    Indent = CInt(modUtils.JsonGetNumber(sJson, "indent", 0))
    Orientation = CInt(modUtils.JsonGetNumber(sJson, "orientation", 0))
    BorderTopStyle = CLng(modUtils.JsonGetNumber(sJson, "borderTopStyle", xlNone))
    BorderTopColor = CLng(modUtils.JsonGetNumber(sJson, "borderTopColor", 0))
    BorderBottomStyle = CLng(modUtils.JsonGetNumber(sJson, "borderBottomStyle", xlNone))
    BorderBottomColor = CLng(modUtils.JsonGetNumber(sJson, "borderBottomColor", 0))
    BorderLeftStyle = CLng(modUtils.JsonGetNumber(sJson, "borderLeftStyle", xlNone))
    BorderLeftColor = CLng(modUtils.JsonGetNumber(sJson, "borderLeftColor", 0))
    BorderRightStyle = CLng(modUtils.JsonGetNumber(sJson, "borderRightStyle", xlNone))
    BorderRightColor = CLng(modUtils.JsonGetNumber(sJson, "borderRightColor", 0))
    MergeRange = modUtils.JsonGetString(sJson, "mergeRange", "")
    ColumnWidth = CSng(modUtils.JsonGetNumber(sJson, "columnWidth", -1))
    RowHeight = CSng(modUtils.JsonGetNumber(sJson, "rowHeight", -1))
    CellLocked = modUtils.JsonGetBool(sJson, "cellLocked", False)
    ConditionalFormats = modUtils.JsonGetRawValue(sJson, "conditionalFormats", "[]")
End Sub

' ===========================================================
'  Create a header-style default format
' ===========================================================
Public Sub SetHeaderDefaults(Optional sFillHex As String = "#003366", _
                              Optional sFontHex As String = "#FFFFFF")
    FontName = "Calibri"
    FontSize = 12
    FontColor = modUtils.HexToLong(sFontHex)
    FontBold = True
    FillColor = modUtils.HexToLong(sFillHex)
    PatternType = xlSolid
    HAlign = xlCenter
    VAlign = xlCenter
    BorderBottomStyle = xlMedium
    BorderBottomColor = 0
    CellLocked = True
End Sub

' ===========================================================
'  Create a currency body-cell format
' ===========================================================
Public Sub SetCurrencyDefaults()
    NumberFormat = "$#,##0.00"
    HAlign = xlRight
    VAlign = xlCenter
End Sub

' ===========================================================
'  Create an alternating row shading format
' ===========================================================
Public Sub SetAlternateRowDefaults(bDark As Boolean)
    If bDark Then
        FillColor = 15921906  ' #F2F2F2 light gray
        PatternType = xlSolid
    Else
        FillColor = 16777215  ' white
        PatternType = xlSolid
    End If
End Sub

' ===========================================================
'  Helper: apply this format to a given Range object
' ===========================================================
Public Sub ApplyToRange(rng As Range)
    On Error GoTo ApplyErr

    ' Font
    With rng.Font
        If FontName <> "" Then .Name = FontName
        If FontSize > 0 Then .Size = FontSize
        .Color = FontColor
        .Bold = FontBold
        .Italic = FontItalic
        .Underline = IIf(FontUnderline, xlUnderlineStyleSingle, xlUnderlineStyleNone)
        .Strikethrough = FontStrikethrough
    End With

    ' Fill
    If FillColor = -1 Then
        rng.Interior.Pattern = xlNone
    Else
        rng.Interior.Color = FillColor
        rng.Interior.Pattern = IIf(PatternType = 0, xlSolid, PatternType)
    End If

    ' Number Format
    If NumberFormat <> "" Then rng.NumberFormat = NumberFormat

    ' Alignment
    rng.HorizontalAlignment = HAlign
    rng.VerticalAlignment = VAlign
    rng.WrapText = WrapText
    rng.IndentLevel = Indent
    If Orientation >= -90 And Orientation <= 90 Then rng.Orientation = Orientation

    ' Borders
    If BorderTopStyle <> xlNone Then
        rng.Borders(xlEdgeTop).LineStyle = BorderTopStyle
        rng.Borders(xlEdgeTop).Color = BorderTopColor
    End If
    If BorderBottomStyle <> xlNone Then
        rng.Borders(xlEdgeBottom).LineStyle = BorderBottomStyle
        rng.Borders(xlEdgeBottom).Color = BorderBottomColor
    End If
    If BorderLeftStyle <> xlNone Then
        rng.Borders(xlEdgeLeft).LineStyle = BorderLeftStyle
        rng.Borders(xlEdgeLeft).Color = BorderLeftColor
    End If
    If BorderRightStyle <> xlNone Then
        rng.Borders(xlEdgeRight).LineStyle = BorderRightStyle
        rng.Borders(xlEdgeRight).Color = BorderRightColor
    End If

    ' Merge
    If MergeRange <> "" Then
        On Error Resume Next
        rng.Worksheet.Range(MergeRange).Merge
        On Error GoTo ApplyErr
    End If

    ' Column width
    If ColumnWidth = 0 Then
        rng.EntireColumn.AutoFit
    ElseIf ColumnWidth > 0 Then
        rng.EntireColumn.ColumnWidth = ColumnWidth
    End If

    ' Row height
    If RowHeight = 0 Then
        rng.EntireRow.AutoFit
    ElseIf RowHeight > 0 Then
        rng.EntireRow.RowHeight = RowHeight
    End If

    ' Lock
    rng.Locked = CellLocked

    Exit Sub
ApplyErr:
    ' Non-fatal — log and continue
    modUtils.WriteAuditLog "WARNING", "clsCellFormat.ApplyToRange", _
        "Failed applying format to " & rng.Address & ": " & Err.Description
    Err.Clear
End Sub

' ===========================================================
'  Private helpers
' ===========================================================
Private Function Escape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    Escape = s
End Function

Private Function BoolToJson(b As Boolean) As String
    BoolToJson = IIf(b, "true", "false")
End Function
