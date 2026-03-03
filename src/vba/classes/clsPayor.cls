VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  clsPayor — Payor configuration object
'  Holds everything the engine needs to process one payor's
'  remittance files: identity, parse rules, lookup config,
'  print settings, and notes.
' ============================================================

' ---- Identity ----------------------------------------------
Public PayorId As String           ' Unique GUID-style ID
Public PayorName As String         ' Display name
Public Aliases() As String         ' Alternative names / header identifiers
Public AliasCount As Integer
Public ExpectedFileType As String  ' xlsx / xls_genuine / xls_html / csv

' ---- Parse Rules -------------------------------------------
' Invoice layout in remittance
Public InvoiceLayout As String     ' "SINGLE_CELL" or "ONE_PER_ROW"
Public InvoiceDelimiter As String  ' Delimiter when SINGLE_CELL
Public InvoiceRegex As String      ' Regex pattern to extract invoice numbers
Public InvoiceColumn As Integer    ' Column index when ONE_PER_ROW (1-based)
Public AmountColumn As Integer     ' Column index for amounts
Public AmountDelimiter As String   ' Delimiter for amounts in SINGLE_CELL mode
Public HeaderRowCount As Integer   ' Rows to skip at top of remittance file
Public InvoiceRow As Integer       ' Row that contains delimited invoice list (SINGLE_CELL)
Public AmountRow As Integer        ' Row for amounts in SINGLE_CELL mode

' ---- Duplicate Handling ------------------------------------
Public DuplicateAction As String   ' "FLAG" | "SKIP" | "INCLUDE"

' ---- Lookup Configuration ----------------------------------
' Stored as JSON; parsed via modDataLookup
Public LookupConfigJson As String

' ---- Print Settings ----------------------------------------
' Stored as JSON; parsed via modFormatting
Public PrintSettingsJson As String

' ---- Metadata ----------------------------------------------
Public Notes As String
Public LastModified As Date
Public CreatedDate As Date

' ===========================================================
'  Initialize with defaults
' ===========================================================
Private Sub Class_Initialize()
    PayorId = ""
    PayorName = ""
    AliasCount = 0
    ReDim Aliases(0)
    ExpectedFileType = FILETYPE_XLSX

    InvoiceLayout = "ONE_PER_ROW"
    InvoiceDelimiter = DELIM_COMMA
    InvoiceRegex = ""
    InvoiceColumn = 1
    AmountColumn = 2
    AmountDelimiter = DELIM_COMMA
    HeaderRowCount = 0
    InvoiceRow = 1
    AmountRow = 1

    DuplicateAction = "FLAG"

    LookupConfigJson = "{}"
    PrintSettingsJson = PrintSettingsDefault()

    Notes = ""
    LastModified = Now()
    CreatedDate = Now()
End Sub

' ===========================================================
'  Add an alias
' ===========================================================
Public Sub AddAlias(sAlias As String)
    ReDim Preserve Aliases(AliasCount)
    Aliases(AliasCount) = Trim(sAlias)
    AliasCount = AliasCount + 1
End Sub

' ===========================================================
'  Serialize to JSON (for storage in PayorRegistry sheet)
' ===========================================================
Public Function ToJson() As String
    Dim j As String
    j = "{"
    j = j & """payorId"":""" & Esc(PayorId) & ""","
    j = j & """payorName"":""" & Esc(PayorName) & ""","
    j = j & """aliases"":" & AliasesToJson() & ","
    j = j & """expectedFileType"":""" & Esc(ExpectedFileType) & ""","
    j = j & """invoiceLayout"":""" & Esc(InvoiceLayout) & ""","
    j = j & """invoiceDelimiter"":""" & Esc(InvoiceDelimiter) & ""","
    j = j & """invoiceRegex"":""" & Esc(InvoiceRegex) & ""","
    j = j & """invoiceColumn"":" & CStr(InvoiceColumn) & ","
    j = j & """amountColumn"":" & CStr(AmountColumn) & ","
    j = j & """amountDelimiter"":""" & Esc(AmountDelimiter) & ""","
    j = j & """headerRowCount"":" & CStr(HeaderRowCount) & ","
    j = j & """invoiceRow"":" & CStr(InvoiceRow) & ","
    j = j & """amountRow"":" & CStr(AmountRow) & ","
    j = j & """duplicateAction"":""" & Esc(DuplicateAction) & ""","
    j = j & """lookupConfig"":" & LookupConfigJson & ","
    j = j & """printSettings"":" & PrintSettingsJson & ","
    j = j & """notes"":""" & Esc(Notes) & ""","
    j = j & """lastModified"":""" & Format(LastModified, "yyyy-mm-dd hh:mm:ss") & ""","
    j = j & """createdDate"":""" & Format(CreatedDate, "yyyy-mm-dd hh:mm:ss") & """"
    j = j & "}"
    ToJson = j
End Function

' ===========================================================
'  Populate from JSON string
' ===========================================================
Public Sub FromJson(sJson As String)
    PayorId = modUtils.JsonGetString(sJson, "payorId", "")
    PayorName = modUtils.JsonGetString(sJson, "payorName", "")
    ExpectedFileType = modUtils.JsonGetString(sJson, "expectedFileType", FILETYPE_XLSX)
    InvoiceLayout = modUtils.JsonGetString(sJson, "invoiceLayout", "ONE_PER_ROW")
    InvoiceDelimiter = modUtils.JsonGetString(sJson, "invoiceDelimiter", DELIM_COMMA)
    InvoiceRegex = modUtils.JsonGetString(sJson, "invoiceRegex", "")
    InvoiceColumn = CInt(modUtils.JsonGetNumber(sJson, "invoiceColumn", 1))
    AmountColumn = CInt(modUtils.JsonGetNumber(sJson, "amountColumn", 2))
    AmountDelimiter = modUtils.JsonGetString(sJson, "amountDelimiter", DELIM_COMMA)
    HeaderRowCount = CInt(modUtils.JsonGetNumber(sJson, "headerRowCount", 0))
    InvoiceRow = CInt(modUtils.JsonGetNumber(sJson, "invoiceRow", 1))
    AmountRow = CInt(modUtils.JsonGetNumber(sJson, "amountRow", 1))
    DuplicateAction = modUtils.JsonGetString(sJson, "duplicateAction", "FLAG")
    LookupConfigJson = modUtils.JsonGetRawValue(sJson, "lookupConfig", "{}")
    PrintSettingsJson = modUtils.JsonGetRawValue(sJson, "printSettings", PrintSettingsDefault())
    Notes = modUtils.JsonGetString(sJson, "notes", "")

    Dim sLM As String: sLM = modUtils.JsonGetString(sJson, "lastModified", "")
    If sLM <> "" Then On Error Resume Next: LastModified = CDate(sLM): On Error GoTo 0

    Dim sCR As String: sCR = modUtils.JsonGetString(sJson, "createdDate", "")
    If sCR <> "" Then On Error Resume Next: CreatedDate = CDate(sCR): On Error GoTo 0

    ' Parse aliases array (simple split by comma within brackets)
    Dim sAliasRaw As String
    sAliasRaw = modUtils.JsonGetRawValue(sJson, "aliases", "[]")
    ParseAliasesFromJson sAliasRaw
End Sub

' ===========================================================
'  Validate — returns "" if valid, else error description
' ===========================================================
Public Function Validate() As String
    If Trim(PayorName) = "" Then
        Validate = "Payor name is required."
        Exit Function
    End If
    If InvoiceLayout = "SINGLE_CELL" And InvoiceDelimiter = "" And InvoiceRegex = "" Then
        Validate = "Single-cell layout requires a delimiter or regex pattern."
        Exit Function
    End If
    If InvoiceLayout = "ONE_PER_ROW" And InvoiceColumn < 1 Then
        Validate = "Invoice column must be >= 1 for one-per-row layout."
        Exit Function
    End If
    If InvoiceRegex <> "" Then
        If Not modUtils.IsValidRegex(InvoiceRegex) Then
            Validate = "Invoice regex pattern is not valid: " & InvoiceRegex
            Exit Function
        End If
    End If
    Validate = ""
End Function

' ===========================================================
'  Get lookup config as parsed Dictionary (via modUtils)
' ===========================================================
Public Function GetLookupConfig() As Object
    Set GetLookupConfig = modUtils.SimpleJsonToDictionary(LookupConfigJson)
End Function

' ===========================================================
'  Get print settings as parsed Dictionary
' ===========================================================
Public Function GetPrintSettings() As Object
    Set GetPrintSettings = modUtils.SimpleJsonToDictionary(PrintSettingsJson)
End Function

' ===========================================================
'  Private helpers
' ===========================================================
Private Function Esc(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(10), "\n")
    Esc = s
End Function

Private Function AliasesToJson() As String
    Dim j As String: j = "["
    Dim i As Integer
    For i = 0 To AliasCount - 1
        If i > 0 Then j = j & ","
        j = j & """" & Esc(Aliases(i)) & """"
    Next i
    j = j & "]"
    AliasesToJson = j
End Function

Private Sub ParseAliasesFromJson(sArr As String)
    ' Very simple: strip [ ], split by ","
    AliasCount = 0
    ReDim Aliases(0)
    sArr = Trim(sArr)
    If Left(sArr, 1) = "[" Then sArr = Mid(sArr, 2)
    If Right(sArr, 1) = "]" Then sArr = Left(sArr, Len(sArr) - 1)
    sArr = Trim(sArr)
    If sArr = "" Then Exit Sub
    Dim parts() As String: parts = Split(sArr, ",")
    Dim p As String
    Dim i As Integer
    For i = 0 To UBound(parts)
        p = Trim(parts(i))
        p = Replace(p, """", "")
        p = Replace(p, "'", "")
        If p <> "" Then
            AddAlias p
        End If
    Next i
End Sub

Private Function PrintSettingsDefault() As String
    PrintSettingsDefault = "{""orientation"":""portrait"",""marginTop"":0.75,""marginBottom"":0.75," & _
        """marginLeft"":0.7,""marginRight"":0.7,""fitToPage"":true,""fitToWidth"":1,""fitToHeight"":0," & _
        """centerHorizontal"":true,""centerVertical"":false,""freezeRow"":1,""freezeCol"":0," & _
        """headerText"":"""",""footerText"":""Page &P of &N""}"
End Function
