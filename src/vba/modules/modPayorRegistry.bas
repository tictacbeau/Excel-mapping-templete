Attribute VB_Name = "modPayorRegistry"
Option Explicit

' ============================================================
'  modPayorRegistry — CRUD for payor configurations
'
'  Storage: Two hidden sheets in ThisWorkbook:
'    SHEET_PAYOR_REGISTRY — one row per payor (identity/rules/config)
'    SHEET_TEMPLATE_DB    — one row per cell in each template
'
'  All data stored as JSON in individual cells.
'  Supports: Add, Get, Update, Delete, Duplicate, List
' ============================================================

' ===========================================================
'  INITIALIZATION — ensure system sheets exist
' ===========================================================
Public Sub InitializeRegistry()
    EnsurePayorRegistrySheet
    EnsureTemplateDbSheet
End Sub

Private Sub EnsurePayorRegistrySheet()
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_PAYOR_REGISTRY, True)

    If ws.Cells(REG_HEADER_ROW, REG_COL_ID).Value = "" Then
        ws.Cells(REG_HEADER_ROW, REG_COL_ID).Value = "PayorId"
        ws.Cells(REG_HEADER_ROW, REG_COL_NAME).Value = "PayorName"
        ws.Cells(REG_HEADER_ROW, REG_COL_ALIASES).Value = "Aliases"
        ws.Cells(REG_HEADER_ROW, REG_COL_FILE_TYPE).Value = "FileType"
        ws.Cells(REG_HEADER_ROW, REG_COL_PARSE_RULES).Value = "ParseRules"
        ws.Cells(REG_HEADER_ROW, REG_COL_LOOKUP_CONFIG).Value = "LookupConfig"
        ws.Cells(REG_HEADER_ROW, REG_COL_PRINT_SETTINGS).Value = "PrintSettings"
        ws.Cells(REG_HEADER_ROW, REG_COL_NOTES).Value = "Notes"
        ws.Cells(REG_HEADER_ROW, REG_COL_LAST_MODIFIED).Value = "LastModified"
        ws.Cells(REG_HEADER_ROW, REG_COL_CREATED).Value = "Created"
        ws.Rows(REG_HEADER_ROW).Font.Bold = True
        ws.Rows(REG_HEADER_ROW).Interior.Color = UI_COLOR_PRIMARY
        ws.Rows(REG_HEADER_ROW).Font.Color = RGB(255, 255, 255)
    End If
End Sub

Private Sub EnsureTemplateDbSheet()
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_TEMPLATE_DB, True)

    If ws.Cells(TDB_HEADER_ROW, TDB_COL_PAYOR_ID).Value = "" Then
        ws.Cells(TDB_HEADER_ROW, TDB_COL_PAYOR_ID).Value = "PayorId"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_TEMPLATE_TYPE).Value = "TemplateType"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_SECTION).Value = "Section"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_ROW_OFFSET).Value = "RowOffset"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_COL_NUM).Value = "ColNum"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_CELL_TYPE).Value = "CellType"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_FIELD_MAP).Value = "FieldMap"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_DEFAULT_VALUE).Value = "DefaultValue"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_FORMAT_JSON).Value = "FormatJson"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_IS_DYNAMIC).Value = "IsDynamic"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_MERGE_REF).Value = "MergeRef"
        ws.Cells(TDB_HEADER_ROW, TDB_COL_FORMULA).Value = "Formula"
        ws.Rows(TDB_HEADER_ROW).Font.Bold = True
        ws.Rows(TDB_HEADER_ROW).Interior.Color = UI_COLOR_PRIMARY
        ws.Rows(TDB_HEADER_ROW).Font.Color = RGB(255, 255, 255)
    End If
End Sub

' ===========================================================
'  GET ALL PAYORS — returns Collection of clsPayor objects
' ===========================================================
Public Function GetAllPayors() As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_PAYOR_REGISTRY, True)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, REG_COL_ID).End(xlUp).Row
    If lastRow < REG_DATA_START_ROW Then
        Set GetAllPayors = col
        Exit Function
    End If

    Dim i As Long
    For i = REG_DATA_START_ROW To lastRow
        Dim sId As String
        sId = Trim(ws.Cells(i, REG_COL_ID).Value)
        If sId <> "" Then
            Dim oPayor As New clsPayor
            RowToPayor ws, i, oPayor
            col.Add oPayor, sId
        End If
    Next i

    Set GetAllPayors = col
End Function

' ===========================================================
'  GET ONE PAYOR by ID — returns clsPayor or Nothing
' ===========================================================
Public Function GetPayor(sPayorId As String) As clsPayor
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_PAYOR_REGISTRY, True)

    Dim nRow As Long
    nRow = FindPayorRow(ws, sPayorId)
    If nRow = 0 Then
        Set GetPayor = Nothing
        Exit Function
    End If

    Dim oPayor As New clsPayor
    RowToPayor ws, nRow, oPayor
    Set GetPayor = oPayor
End Function

' ===========================================================
'  GET PAYOR BY NAME — returns clsPayor or Nothing
' ===========================================================
Public Function GetPayorByName(sName As String) As clsPayor
    Dim col As Collection
    Set col = GetAllPayors()
    Dim oPayor As clsPayor
    For Each oPayor In col
        If LCase(oPayor.PayorName) = LCase(sName) Then
            Set GetPayorByName = oPayor
            Exit Function
        End If
        ' Check aliases
        Dim j As Integer
        For j = 0 To oPayor.AliasCount - 1
            If LCase(oPayor.Aliases(j)) = LCase(sName) Then
                Set GetPayorByName = oPayor
                Exit Function
            End If
        Next j
    Next oPayor
    Set GetPayorByName = Nothing
End Function

' ===========================================================
'  SAVE PAYOR — create or update
' ===========================================================
Public Function SavePayor(oPayor As clsPayor) As Boolean
    On Error GoTo SaveErr
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_PAYOR_REGISTRY, True)

    ' Generate ID if new
    If oPayor.PayorId = "" Then
        oPayor.PayorId = modUtils.GenerateId()
        oPayor.CreatedDate = Now()
    End If
    oPayor.LastModified = Now()

    Dim nRow As Long
    nRow = FindPayorRow(ws, oPayor.PayorId)
    If nRow = 0 Then
        ' New row
        nRow = ws.Cells(ws.Rows.Count, REG_COL_ID).End(xlUp).Row + 1
        If nRow < REG_DATA_START_ROW Then nRow = REG_DATA_START_ROW
    End If

    PayorToRow oPayor, ws, nRow

    modUtils.WriteAuditLog "SAVE_PAYOR", "PayorRegistry", _
        "Saved payor: " & oPayor.PayorName, oPayor.PayorName, "", "OK"
    SavePayor = True
    Exit Function

SaveErr:
    modUtils.WriteAuditLog "ERROR", "PayorRegistry.SavePayor", _
        Err.Description, oPayor.PayorName, "", "FAIL"
    SavePayor = False
    Err.Clear
End Function

' ===========================================================
'  DELETE PAYOR by ID
' ===========================================================
Public Function DeletePayor(sPayorId As String) As Boolean
    On Error GoTo DelErr
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_PAYOR_REGISTRY, True)

    Dim nRow As Long
    nRow = FindPayorRow(ws, sPayorId)
    If nRow = 0 Then
        DeletePayor = False
        Exit Function
    End If

    Dim sName As String
    sName = ws.Cells(nRow, REG_COL_NAME).Value
    ws.Rows(nRow).Delete

    ' Also delete template rows
    DeletePayorTemplates sPayorId

    modUtils.WriteAuditLog "DELETE_PAYOR", "PayorRegistry", _
        "Deleted payor: " & sName & " (" & sPayorId & ")", sName, "", "OK"
    DeletePayor = True
    Exit Function

DelErr:
    DeletePayor = False
    Err.Clear
End Function

' ===========================================================
'  DUPLICATE PAYOR — clone config, assign new ID/name
' ===========================================================
Public Function DuplicatePayor(sSourceId As String, sNewName As String) As clsPayor
    Dim oSource As clsPayor
    Set oSource = GetPayor(sSourceId)
    If oSource Is Nothing Then
        Set DuplicatePayor = Nothing
        Exit Function
    End If

    ' Clone by JSON round-trip
    Dim oNew As New clsPayor
    Dim sJson As String
    sJson = BuildPayorJson(oSource)
    LoadPayorFromJson oNew, sJson

    oNew.PayorId = ""  ' Will be assigned on save
    oNew.PayorName = sNewName
    oNew.CreatedDate = Now()
    oNew.LastModified = Now()

    ' Save first so oNew.PayorId is populated before template duplication
    If SavePayor(oNew) Then
        DuplicatePayorTemplates sSourceId, oNew
        Set DuplicatePayor = oNew
    Else
        Set DuplicatePayor = Nothing
    End If
End Function

' ===========================================================
'  PAYOR NAME EXISTS CHECK
' ===========================================================
Public Function PayorNameExists(sName As String, _
                                 Optional sExcludeId As String = "") As Boolean
    Dim col As Collection
    Set col = GetAllPayors()
    Dim oPayor As clsPayor
    For Each oPayor In col
        If LCase(oPayor.PayorName) = LCase(sName) Then
            If oPayor.PayorId <> sExcludeId Then
                PayorNameExists = True
                Exit Function
            End If
        End If
    Next oPayor
    PayorNameExists = False
End Function

' ===========================================================
'  GET PAYOR NAMES for dropdowns (returns sorted String array)
' ===========================================================
Public Function GetPayorNameList() As String()
    Dim col As Collection
    Set col = GetAllPayors()
    Dim result() As String
    Dim count As Integer: count = col.Count
    If count = 0 Then
        ReDim result(0): result(0) = ""
        GetPayorNameList = result
        Exit Function
    End If
    ReDim result(count - 1)
    Dim i As Integer: i = 0
    Dim oPayor As clsPayor
    For Each oPayor In col
        result(i) = oPayor.PayorName
        i = i + 1
    Next oPayor
    ' Simple bubble sort
    Dim j As Integer, temp As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If result(i) > result(j) Then
                temp = result(i): result(i) = result(j): result(j) = temp
            End If
        Next j
    Next i
    GetPayorNameList = result
End Function

' ===========================================================
'  TEMPLATE CRUD
' ===========================================================

' Save template (all cells) for a payor
Public Function SaveTemplate(oTemplate As clsTemplate) As Boolean
    On Error GoTo TmplErr
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_TEMPLATE_DB, True)

    ' Delete existing rows for this payor+type
    DeleteTemplateRows ws, oTemplate.PayorId, oTemplate.TemplateType

    ' Write each cell definition
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TDB_COL_PAYOR_ID).End(xlUp).Row + 1
    If lastRow < TDB_DATA_START_ROW Then lastRow = TDB_DATA_START_ROW

    Dim i As Integer
    For i = 0 To oTemplate.CellCount - 1
        Dim sCell As String: sCell = oTemplate.CellJson(i)
        Dim nRow As Long: nRow = lastRow + i
        ws.Cells(nRow, TDB_COL_PAYOR_ID).Value = oTemplate.PayorId
        ws.Cells(nRow, TDB_COL_TEMPLATE_TYPE).Value = oTemplate.TemplateType
        ws.Cells(nRow, TDB_COL_SECTION).Value = modUtils.JsonGetString(sCell, "section", "")
        ws.Cells(nRow, TDB_COL_ROW_OFFSET).Value = modUtils.JsonGetNumber(sCell, "rowOffset", 1)
        ws.Cells(nRow, TDB_COL_COL_NUM).Value = modUtils.JsonGetNumber(sCell, "col", 1)
        ws.Cells(nRow, TDB_COL_CELL_TYPE).Value = modUtils.JsonGetString(sCell, "cellType", CELLTYPE_LABEL)
        ws.Cells(nRow, TDB_COL_FIELD_MAP).Value = modUtils.JsonGetString(sCell, "fieldMap", "")
        ws.Cells(nRow, TDB_COL_DEFAULT_VALUE).Value = modUtils.JsonGetString(sCell, "defaultValue", "")
        ws.Cells(nRow, TDB_COL_FORMAT_JSON).Value = modUtils.JsonGetRawValue(sCell, "formatJson", "{}")
        ws.Cells(nRow, TDB_COL_IS_DYNAMIC).Value = modUtils.JsonGetBool(sCell, "isDynamic", False)
        ws.Cells(nRow, TDB_COL_MERGE_REF).Value = modUtils.JsonGetString(sCell, "mergeRef", "")
        ws.Cells(nRow, TDB_COL_FORMULA).Value = modUtils.JsonGetString(sCell, "formula", "")
    Next i

    ' Also save template metadata in payor row
    Dim oPayor As clsPayor
    Set oPayor = GetPayor(oTemplate.PayorId)
    If Not oPayor Is Nothing Then
        ' Store template metadata in payor (header/body/footer counts, etc.)
        ' This is lightweight - full cell data is in TemplateDB
        SavePayor oPayor  ' refresh last-modified
    End If

    SaveTemplate = True
    Exit Function

TmplErr:
    modUtils.WriteAuditLog "ERROR", "PayorRegistry.SaveTemplate", Err.Description, , "", "FAIL"
    SaveTemplate = False
    Err.Clear
End Function

' Load template for a payor (returns clsTemplate or Nothing)
Public Function LoadTemplate(sPayorId As String, sTemplateType As String) As clsTemplate
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_TEMPLATE_DB, True)

    Dim oTemplate As New clsTemplate
    oTemplate.PayorId = sPayorId
    oTemplate.TemplateType = sTemplateType

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TDB_COL_PAYOR_ID).End(xlUp).Row
    If lastRow < TDB_DATA_START_ROW Then
        Set LoadTemplate = Nothing
        Exit Function
    End If

    Dim found As Boolean: found = False
    Dim i As Long
    For i = TDB_DATA_START_ROW To lastRow
        Dim sId As String: sId = Trim(ws.Cells(i, TDB_COL_PAYOR_ID).Value)
        Dim sType As String: sType = Trim(ws.Cells(i, TDB_COL_TEMPLATE_TYPE).Value)
        If sId = sPayorId And sType = sTemplateType Then
            found = True
            ' Build cell JSON from row
            Dim cellJson As String
            cellJson = BuildCellJson(ws, i)
            oTemplate.AddCell _
                modUtils.JsonGetString(cellJson, "section", SECTION_HEADER), _
                CInt(modUtils.JsonGetNumber(cellJson, "rowOffset", 1)), _
                CInt(modUtils.JsonGetNumber(cellJson, "col", 1)), _
                modUtils.JsonGetString(cellJson, "cellType", CELLTYPE_LABEL), _
                modUtils.JsonGetString(cellJson, "fieldMap", ""), _
                modUtils.JsonGetString(cellJson, "defaultValue", ""), _
                modUtils.JsonGetRawValue(cellJson, "formatJson", "{}"), _
                modUtils.JsonGetBool(cellJson, "isDynamic", False), _
                modUtils.JsonGetString(cellJson, "mergeRef", ""), _
                modUtils.JsonGetString(cellJson, "formula", "")
        End If
    Next i

    If Not found Then
        Set LoadTemplate = Nothing
    Else
        Set LoadTemplate = oTemplate
    End If
End Function

' Delete template rows for a payor+type
Private Sub DeleteTemplateRows(ws As Worksheet, sPayorId As String, sType As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TDB_COL_PAYOR_ID).End(xlUp).Row
    Dim i As Long
    ' Delete from bottom up
    For i = lastRow To TDB_DATA_START_ROW Step -1
        If Trim(ws.Cells(i, TDB_COL_PAYOR_ID).Value) = sPayorId And _
           Trim(ws.Cells(i, TDB_COL_TEMPLATE_TYPE).Value) = sType Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Private Sub DeletePayorTemplates(sPayorId As String)
    Dim ws As Worksheet
    Set ws = modUtils.GetOrCreateSheet(SHEET_TEMPLATE_DB, True)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, TDB_COL_PAYOR_ID).End(xlUp).Row
    Dim i As Long
    For i = lastRow To TDB_DATA_START_ROW Step -1
        If Trim(ws.Cells(i, TDB_COL_PAYOR_ID).Value) = sPayorId Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Private Sub DuplicatePayorTemplates(sSourceId As String, oNewPayor As clsPayor)
    ' The new payor ID won't exist yet — store source templates under new payor
    ' after save, the ID will be assigned
    Dim tmplTypes(1) As String
    tmplTypes(0) = TMPL_SINGLE
    tmplTypes(1) = TMPL_MULTI
    Dim t As Integer
    For t = 0 To 1
        Dim oTmpl As clsTemplate
        Set oTmpl = LoadTemplate(sSourceId, tmplTypes(t))
        If Not oTmpl Is Nothing Then
            oTmpl.PayorId = oNewPayor.PayorId
            SaveTemplate oTmpl
        End If
    Next t
End Sub

' ===========================================================
'  PRIVATE HELPERS
' ===========================================================

' Find the row number for a payor ID (0 = not found)
Private Function FindPayorRow(ws As Worksheet, sPayorId As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, REG_COL_ID).End(xlUp).Row
    Dim i As Long
    For i = REG_DATA_START_ROW To lastRow
        If Trim(ws.Cells(i, REG_COL_ID).Value) = sPayorId Then
            FindPayorRow = i
            Exit Function
        End If
    Next i
    FindPayorRow = 0
End Function

' Read a row from PayorRegistry sheet into a clsPayor object
Private Sub RowToPayor(ws As Worksheet, nRow As Long, oPayor As clsPayor)
    oPayor.PayorId = Trim(ws.Cells(nRow, REG_COL_ID).Value)
    oPayor.PayorName = Trim(ws.Cells(nRow, REG_COL_NAME).Value)
    oPayor.ExpectedFileType = Trim(ws.Cells(nRow, REG_COL_FILE_TYPE).Value)
    oPayor.Notes = Trim(ws.Cells(nRow, REG_COL_NOTES).Value)
    oPayor.LookupConfigJson = ws.Cells(nRow, REG_COL_LOOKUP_CONFIG).Value
    oPayor.PrintSettingsJson = ws.Cells(nRow, REG_COL_PRINT_SETTINGS).Value

    On Error Resume Next
    oPayor.LastModified = CDate(ws.Cells(nRow, REG_COL_LAST_MODIFIED).Value)
    oPayor.CreatedDate = CDate(ws.Cells(nRow, REG_COL_CREATED).Value)
    On Error GoTo 0

    ' Parse rules from stored JSON
    Dim sRules As String
    sRules = ws.Cells(nRow, REG_COL_PARSE_RULES).Value
    If sRules <> "" Then
        oPayor.InvoiceLayout = modUtils.JsonGetString(sRules, "invoiceLayout", "ONE_PER_ROW")
        oPayor.InvoiceDelimiter = modUtils.JsonGetString(sRules, "invoiceDelimiter", DELIM_COMMA)
        oPayor.InvoiceRegex = modUtils.JsonGetString(sRules, "invoiceRegex", "")
        oPayor.InvoiceColumn = CInt(modUtils.JsonGetNumber(sRules, "invoiceColumn", 1))
        oPayor.AmountColumn = CInt(modUtils.JsonGetNumber(sRules, "amountColumn", 2))
        oPayor.AmountDelimiter = modUtils.JsonGetString(sRules, "amountDelimiter", DELIM_COMMA)
        oPayor.HeaderRowCount = CInt(modUtils.JsonGetNumber(sRules, "headerRowCount", 0))
        oPayor.InvoiceRow = CInt(modUtils.JsonGetNumber(sRules, "invoiceRow", 1))
        oPayor.AmountRow = CInt(modUtils.JsonGetNumber(sRules, "amountRow", 1))
        oPayor.DuplicateAction = modUtils.JsonGetString(sRules, "duplicateAction", "FLAG")
    End If

    ' Parse aliases
    Dim sAliases As String
    sAliases = ws.Cells(nRow, REG_COL_ALIASES).Value
    If sAliases <> "" And sAliases <> "[]" Then
        Dim parts() As String
        parts = Split(Replace(Replace(sAliases, "[", ""), "]", ""), ",")
        Dim i As Integer
        For i = 0 To UBound(parts)
            Dim p As String: p = Trim(Replace(parts(i), """", ""))
            If p <> "" Then oPayor.AddAlias p
        Next i
    End If
End Sub

' Write a clsPayor object to a row in PayorRegistry sheet
Private Sub PayorToRow(oPayor As clsPayor, ws As Worksheet, nRow As Long)
    ws.Cells(nRow, REG_COL_ID).Value = oPayor.PayorId
    ws.Cells(nRow, REG_COL_NAME).Value = oPayor.PayorName
    ws.Cells(nRow, REG_COL_FILE_TYPE).Value = oPayor.ExpectedFileType
    ws.Cells(nRow, REG_COL_NOTES).Value = oPayor.Notes
    ws.Cells(nRow, REG_COL_LOOKUP_CONFIG).Value = oPayor.LookupConfigJson
    ws.Cells(nRow, REG_COL_PRINT_SETTINGS).Value = oPayor.PrintSettingsJson
    ws.Cells(nRow, REG_COL_LAST_MODIFIED).Value = oPayor.LastModified
    ws.Cells(nRow, REG_COL_CREATED).Value = oPayor.CreatedDate

    ' Build parse rules JSON — escape all string values
    Dim sRules As String
    sRules = "{" & _
        """invoiceLayout"":""" & EscJson(oPayor.InvoiceLayout) & """," & _
        """invoiceDelimiter"":""" & EscJson(oPayor.InvoiceDelimiter) & """," & _
        """invoiceRegex"":""" & EscJson(oPayor.InvoiceRegex) & """," & _
        """invoiceColumn"":" & oPayor.InvoiceColumn & "," & _
        """amountColumn"":" & oPayor.AmountColumn & "," & _
        """amountDelimiter"":""" & EscJson(oPayor.AmountDelimiter) & """," & _
        """headerRowCount"":" & oPayor.HeaderRowCount & "," & _
        """invoiceRow"":" & oPayor.InvoiceRow & "," & _
        """amountRow"":" & oPayor.AmountRow & "," & _
        """duplicateAction"":""" & EscJson(oPayor.DuplicateAction) & """" & _
        "}"
    ws.Cells(nRow, REG_COL_PARSE_RULES).Value = sRules

    ' Build aliases JSON array — escape alias strings
    Dim sAliases As String: sAliases = "["
    Dim i As Integer
    For i = 0 To oPayor.AliasCount - 1
        If i > 0 Then sAliases = sAliases & ","
        sAliases = sAliases & """" & EscJson(oPayor.Aliases(i)) & """"
    Next i
    sAliases = sAliases & "]"
    ws.Cells(nRow, REG_COL_ALIASES).Value = sAliases
End Sub

' Build cell JSON from TemplateDB row
Private Function BuildCellJson(ws As Worksheet, nRow As Long) As String
    Dim j As String
    j = "{"
    j = j & """section"":""" & EscJson(ws.Cells(nRow, TDB_COL_SECTION).Value) & ""","
    j = j & """rowOffset"":" & CStr(ws.Cells(nRow, TDB_COL_ROW_OFFSET).Value) & ","
    j = j & """col"":" & CStr(ws.Cells(nRow, TDB_COL_COL_NUM).Value) & ","
    j = j & """cellType"":""" & EscJson(ws.Cells(nRow, TDB_COL_CELL_TYPE).Value) & ""","
    j = j & """fieldMap"":""" & EscJson(ws.Cells(nRow, TDB_COL_FIELD_MAP).Value) & ""","
    j = j & """defaultValue"":""" & EscJson(ws.Cells(nRow, TDB_COL_DEFAULT_VALUE).Value) & ""","
    j = j & """formatJson"":" & IIf(ws.Cells(nRow, TDB_COL_FORMAT_JSON).Value = "", "{}", ws.Cells(nRow, TDB_COL_FORMAT_JSON).Value) & ","
    j = j & """isDynamic"":" & IIf(CBool(ws.Cells(nRow, TDB_COL_IS_DYNAMIC).Value), "true", "false") & ","
    j = j & """mergeRef"":""" & EscJson(ws.Cells(nRow, TDB_COL_MERGE_REF).Value) & ""","
    j = j & """formula"":""" & EscJson(ws.Cells(nRow, TDB_COL_FORMULA).Value) & """"
    j = j & "}"
    BuildCellJson = j
End Function

Private Function EscJson(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    EscJson = s
End Function

Private Function BuildPayorJson(oPayor As clsPayor) As String
    BuildPayorJson = oPayor.ToJson()
End Function

Private Sub LoadPayorFromJson(oPayor As clsPayor, sJson As String)
    oPayor.FromJson sJson
End Sub
