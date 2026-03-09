Attribute VB_Name = "modTemplateEngine"
Option Explicit

' ============================================================
'  modTemplateEngine — Template Expansion & Population
'
'  Core processing engine:
'    1. Parse remittance data using payor rules
'    2. Select appropriate template (SINGLE vs MULTI)
'    3. Expand body section N times (multi-invoice)
'    4. Populate all mapped cells
'    5. Apply all formatting
'    6. Run internal data lookups
'    7. Calculate formula cells
'    8. Set print area / freeze panes
'    9. Flag errors
' ============================================================

' ===========================================================
'  MAIN ENTRY POINT — full processing workflow
'  Returns the output worksheet, or Nothing on failure.
' ===========================================================
Public Function ProcessRemittance(sPayorId As String, _
                                   sSourceFile As String) As Worksheet
    Set ProcessRemittance = Nothing

    ' 1. Load payor config
    Dim oPayor As clsPayor
    Set oPayor = modPayorRegistry.GetPayor(sPayorId)
    If oPayor Is Nothing Then
        modUtils.ShowError "Payor not found: " & sPayorId
        Exit Function
    End If

    ' 2. Initialize remittance object + convert file
    Dim oRemittance As New clsRemittance
    oRemittance.PayorId = sPayorId
    oRemittance.PayorName = oPayor.PayorName
    oRemittance.SourceFilePath = sSourceFile

    Dim sNormalizedPath As String
    sNormalizedPath = modFileConverter.ConvertToXlsx(sSourceFile, oRemittance)
    If sNormalizedPath = "" Then
        modUtils.ShowError "File conversion failed. Check the audit log for details."
        Exit Function
    End If
    oRemittance.NormalizedFilePath = sNormalizedPath

    ' 3. Open normalized file and parse remittance data
    Dim wbSource As Workbook
    On Error GoTo ProcessErr
    Application.DisplayAlerts = False
    Set wbSource = Application.Workbooks.Open(sNormalizedPath, ReadOnly:=True)
    Application.DisplayAlerts = True
    Set oRemittance.RawWorkbook = wbSource

    ParseRemittanceData oPayor, oRemittance, wbSource.Sheets(1)

    If oRemittance.InvoiceCount = 0 Then
        wbSource.Close SaveChanges:=False
        modUtils.ShowError "No invoices could be extracted from the remittance file." & vbCrLf & _
            "Check the payor parse rules and the file format." & vbCrLf & vbCrLf & _
            "Warnings:" & vbCrLf & oRemittance.AllWarnings
        Exit Function
    End If

    ' 4. Select template type
    Dim sTemplateType As String
    sTemplateType = IIf(oRemittance.IsSingleInvoice, TMPL_SINGLE, TMPL_MULTI)
    oRemittance.AddAuditLine "Template type selected: " & sTemplateType & _
        " (" & oRemittance.InvoiceCount & " invoice(s))"

    ' 5. Load template
    Dim oTemplate As clsTemplate
    Set oTemplate = modPayorRegistry.LoadTemplate(sPayorId, sTemplateType)
    If oTemplate Is Nothing Then
        ' Try the other template type as fallback
        Dim sFallback As String
        sFallback = IIf(sTemplateType = TMPL_SINGLE, TMPL_MULTI, TMPL_SINGLE)
        Set oTemplate = modPayorRegistry.LoadTemplate(sPayorId, sFallback)
        If oTemplate Is Nothing Then
            wbSource.Close SaveChanges:=False
            modUtils.ShowError "No template found for payor: " & oPayor.PayorName & vbCrLf & _
                "Please configure templates using the Payor Setup Wizard."
            Exit Function
        End If
        oRemittance.AddAuditLine "Using fallback template type: " & sFallback
    End If

    ' 6. Create output worksheet
    Dim wsOutput As Worksheet
    Set wsOutput = CreateOutputSheet(oPayor.PayorName)

    ' 7. Expand + populate template
    ExpandAndPopulateTemplate oTemplate, oRemittance, oPayor, wsOutput

    ' 8. Run lookups from Data Model
    If oPayor.LookupConfigJson <> "{}" And oPayor.LookupConfigJson <> "" Then
        RunDataLookups oTemplate, oRemittance, oPayor, wsOutput
    End If

    ' 9. Calculate formula cells
    wsOutput.Calculate

    ' 10. Apply print settings
    ApplyPrintSettings wsOutput, oPayor, oTemplate

    ' 11. Flag any issues
    FlagLookupErrors wsOutput, oRemittance
    FlagDuplicates wsOutput, oRemittance, oTemplate

    ' 12. Close source workbook + cleanup temp file
    wbSource.Close SaveChanges:=False
    Set oRemittance.RawWorkbook = Nothing
    On Error Resume Next
    Kill sNormalizedPath
    On Error GoTo 0

    ' 13. Log and return
    modUtils.WriteAuditLog "PROCESS", "TemplateEngine", _
        oRemittance.ToSummaryString(), oPayor.PayorName, sSourceFile, "OK"
    oRemittance.AddAuditLine "Processing complete. Output: " & wsOutput.Name

    wsOutput.Activate
    Set ProcessRemittance = wsOutput
    Exit Function

ProcessErr:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    Kill sNormalizedPath
    On Error GoTo 0
    modUtils.ShowError "Processing error: " & Err.Description
    modUtils.WriteAuditLog "ERROR", "TemplateEngine.ProcessRemittance", _
        Err.Description, sPayorId, sSourceFile, "FAIL"
    Err.Clear
End Function

' ===========================================================
'  PARSE REMITTANCE DATA from normalized worksheet
' ===========================================================
Private Sub ParseRemittanceData(oPayor As clsPayor, oRemittance As clsRemittance, _
                                  ws As Worksheet)
    oRemittance.AddAuditLine "Parsing remittance data. Layout: " & oPayor.InvoiceLayout

    Select Case UCase(oPayor.InvoiceLayout)
        Case "SINGLE_CELL"
            ParseSingleCellLayout oPayor, oRemittance, ws
        Case "ONE_PER_ROW"
            ParseOnePerRowLayout oPayor, oRemittance, ws
        Case Else
            ParseOnePerRowLayout oPayor, oRemittance, ws
    End Select

    oRemittance.AddAuditLine "Parse complete. " & oRemittance.InvoiceCount & " invoice(s) extracted."
End Sub

' Parse: All invoice numbers in one cell (comma/delimiter-separated)
Private Sub ParseSingleCellLayout(oPayor As clsPayor, oRemittance As clsRemittance, _
                                    ws As Worksheet)
    ' Get raw invoice string from configured cell
    Dim nInvRow As Integer: nInvRow = oPayor.InvoiceRow + oPayor.HeaderRowCount
    Dim nInvCol As Integer: nInvCol = oPayor.InvoiceColumn
    Dim nAmtRow As Integer: nAmtRow = oPayor.AmountRow + oPayor.HeaderRowCount
    Dim nAmtCol As Integer: nAmtCol = oPayor.AmountColumn

    Dim sRawInvoices As String
    sRawInvoices = modUtils.CleanCellValue(ws.Cells(nInvRow, nInvCol).Value)
    Dim sRawAmounts As String
    sRawAmounts = modUtils.CleanCellValue(ws.Cells(nAmtRow, nAmtCol).Value)

    oRemittance.AddAuditLine "Raw invoice cell (" & nInvRow & "," & nInvCol & "): " & sRawInvoices
    oRemittance.AddAuditLine "Raw amount cell (" & nAmtRow & "," & nAmtCol & "): " & sRawAmounts

    ' Extract invoice numbers using regex or delimiter
    Dim invoiceNums() As String
    If oPayor.InvoiceRegex <> "" Then
        invoiceNums = modUtils.RegexExtractAll(sRawInvoices, oPayor.InvoiceRegex)
    Else
        invoiceNums = modUtils.SplitDelimited(sRawInvoices, oPayor.InvoiceDelimiter)
    End If

    ' Extract amounts
    Dim amounts() As String
    amounts = modUtils.SplitDelimited(sRawAmounts, oPayor.AmountDelimiter)

    ' Pair up invoices with amounts
    Dim i As Integer
    For i = 0 To UBound(invoiceNums)
        Dim sInv As String: sInv = Trim(invoiceNums(i))
        If sInv = "" Then GoTo NextInv
        Dim dAmt As Double: dAmt = 0
        If i <= UBound(amounts) Then
            dAmt = modUtils.ParseCurrencyString(amounts(i))
        End If
        oRemittance.AddInvoice sInv, dAmt
NextInv:
    Next i

    ' Try to get payment total from the sheet
    ExtractPaymentHeader oPayor, oRemittance, ws
End Sub

' Parse: One invoice per row in configured column
Private Sub ParseOnePerRowLayout(oPayor As clsPayor, oRemittance As clsRemittance, _
                                   ws As Worksheet)
    Dim nStartRow As Long: nStartRow = oPayor.HeaderRowCount + 1
    Dim nLastRow As Long: nLastRow = ws.Cells(ws.Rows.Count, oPayor.InvoiceColumn).End(xlUp).Row
    Dim nInvCol As Integer: nInvCol = oPayor.InvoiceColumn
    Dim nAmtCol As Integer: nAmtCol = oPayor.AmountColumn

    oRemittance.AddAuditLine "Scanning rows " & nStartRow & " to " & nLastRow & _
        " for invoices in column " & nInvCol

    Dim i As Long
    For i = nStartRow To nLastRow
        Dim sRaw As String
        sRaw = modUtils.CleanCellValue(ws.Cells(i, nInvCol).Value)
        If sRaw = "" Then GoTo NextRow

        ' Extract invoice numbers from cell (may contain regex-matched values)
        Dim invoiceNums() As String
        If oPayor.InvoiceRegex <> "" Then
            invoiceNums = modUtils.RegexExtractAll(sRaw, oPayor.InvoiceRegex)
        Else
            ReDim invoiceNums(0): invoiceNums(0) = sRaw
        End If

        Dim dAmt As Double
        dAmt = modUtils.ParseCurrencyString( _
            modUtils.CleanCellValue(ws.Cells(i, nAmtCol).Value))

        Dim j As Integer
        For j = 0 To UBound(invoiceNums)
            Dim sInv As String: sInv = Trim(invoiceNums(j))
            If sInv <> "" Then
                Dim sRowData As String
                sRowData = BuildRowDataJson(ws, i)
                ' Date column is not configured per-payor; pass empty string.
                ' Date values are available via sRowData (raw JSON) if needed.
                oRemittance.AddInvoice sInv, dAmt, "", sRowData
            End If
        Next j
NextRow:
    Next i

    ExtractPaymentHeader oPayor, oRemittance, ws
End Sub

' Try to extract payment date, total, check number from top rows
Private Sub ExtractPaymentHeader(oPayor As clsPayor, oRemittance As clsRemittance, _
                                   ws As Worksheet)
    ' Scan header rows for known payment fields
    Dim i As Long
    Dim maxRow As Long: maxRow = IIf(oPayor.HeaderRowCount > 0, oPayor.HeaderRowCount, 5)

    For i = 1 To maxRow
        Dim j As Integer
        For j = 1 To 10
            Dim sCell As String
            sCell = LCase(Trim(modUtils.CleanCellValue(ws.Cells(i, j).Value)))
            Dim sNext As String
            sNext = modUtils.CleanCellValue(ws.Cells(i, j + 1).Value)

            If InStr(sCell, "date") > 0 And oRemittance.PaymentDate = "" Then
                oRemittance.PaymentDate = sNext
            ElseIf (InStr(sCell, "total") > 0 Or InStr(sCell, "amount") > 0) _
                And oRemittance.PaymentAmount = 0 Then
                Dim d As Double: d = modUtils.ParseCurrencyString(sNext)
                If d <> 0 Then oRemittance.PaymentAmount = d
            ElseIf (InStr(sCell, "check") > 0 Or InStr(sCell, "eft") > 0 Or _
                    InStr(sCell, "reference") > 0) And oRemittance.CheckNumber = "" Then
                oRemittance.CheckNumber = sNext
            End If
        Next j
    Next i

    ' If no total found, set to sum of extracted invoices
    If oRemittance.PaymentAmount = 0 Then
        oRemittance.PaymentAmount = oRemittance.TotalExtracted
    End If
End Sub

' Build a simple JSON object from a row's cell values
Private Function BuildRowDataJson(ws As Worksheet, nRow As Long) As String
    Dim j As String: j = "{"
    Dim nLastCol As Integer
    nLastCol = ws.Cells(nRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Integer
    For c = 1 To IIf(nLastCol > 20, 20, nLastCol)
        If c > 1 Then j = j & ","
        j = j & """c" & c & """:""" & _
            Replace(modUtils.CleanCellValue(ws.Cells(nRow, c).Value), """", "\""") & """"
    Next c
    j = j & "}"
    BuildRowDataJson = j
End Function

' ===========================================================
'  EXPAND AND POPULATE TEMPLATE
' ===========================================================
Private Sub ExpandAndPopulateTemplate(oTemplate As clsTemplate, _
                                        oRemittance As clsRemittance, _
                                        oPayor As clsPayor, _
                                        wsOut As Worksheet)
    wsOut.Cells.Clear

    ' Determine layout
    Dim nInvoices As Integer: nInvoices = oRemittance.InvoiceCount
    Dim nHeaderRows As Integer: nHeaderRows = oTemplate.HeaderRowCount
    Dim nBodyRowsPerInvoice As Integer: nBodyRowsPerInvoice = oTemplate.BodyRowCount
    Dim nFooterRows As Integer: nFooterRows = oTemplate.FooterRowCount

    ' Calculate absolute row positions
    Dim nBodyStart As Long: nBodyStart = nHeaderRows + 1
    Dim nBodyEnd As Long: nBodyEnd = nBodyStart + (nBodyRowsPerInvoice * nInvoices) - 1
    Dim nFooterStart As Long: nFooterStart = nBodyEnd + 1

    oTemplate.BodyStartRow = nBodyStart
    oTemplate.BodyEndRow = nBodyEnd

    ' --- Populate HEADER cells ---
    Dim cells As Collection
    Set cells = oTemplate.GetSectionCells(SECTION_HEADER)
    Dim sCell As Variant
    For Each sCell In cells
        Dim nRow As Long: nRow = modUtils.JsonGetNumber(sCell, "rowOffset", 1)
        Dim nCol As Integer: nCol = CInt(modUtils.JsonGetNumber(sCell, "col", 1))
        PopulateCell wsOut, nRow, nCol, sCell, oRemittance, 0
    Next sCell

    ' --- Populate BODY cells (repeat for each invoice) ---
    Set cells = oTemplate.GetBodyCells()
    Dim iInv As Integer
    For iInv = 0 To nInvoices - 1
        Dim nBaseRow As Long: nBaseRow = nBodyStart + (iInv * nBodyRowsPerInvoice)

        ' Apply alternating row shading
        If oTemplate.AlternateRowShading Then
            Dim oFmt As New clsCellFormat
            oFmt.SetAlternateRowDefaults (iInv Mod 2 = 1)
            If oFmt.FillColor <> -1 Then
                Dim rng As Range
                Set rng = wsOut.Range( _
                    wsOut.Cells(nBaseRow, 1), _
                    wsOut.Cells(nBaseRow + nBodyRowsPerInvoice - 1, oTemplate.ColumnCount))
                rng.Interior.Color = oFmt.FillColor
            End If
        End If

        For Each sCell In cells
            Dim nRowOffset As Long: nRowOffset = modUtils.JsonGetNumber(sCell, "rowOffset", 1)
            Dim nCellCol As Integer: nCellCol = CInt(modUtils.JsonGetNumber(sCell, "col", 1))
            Dim nAbsRow As Long: nAbsRow = nBaseRow + nRowOffset - 1
            PopulateCell wsOut, nAbsRow, nCellCol, sCell, oRemittance, iInv
        Next sCell
    Next iInv

    ' --- Populate FOOTER cells ---
    Set cells = oTemplate.GetSectionCells(SECTION_FOOTER)
    For Each sCell In cells
        Dim nFRow As Long: nFRow = nFooterStart + modUtils.JsonGetNumber(sCell, "rowOffset", 1) - 1
        Dim nFCol As Integer: nFCol = CInt(modUtils.JsonGetNumber(sCell, "col", 1))
        PopulateCell wsOut, nFRow, nFCol, sCell, oRemittance, -1
    Next sCell

    ' Apply all cell-level formatting
    ApplyTemplateFormatting wsOut, oTemplate, nBodyStart, nBodyEnd, nFooterStart, nInvoices

    oRemittance.AddAuditLine "Template expanded: " & nInvoices & " invoice row(s), " & _
        "total rows: " & (nFooterStart + nFooterRows - 1)
End Sub

' ===========================================================
'  POPULATE A SINGLE CELL based on its type and field mapping
' ===========================================================
Private Sub PopulateCell(wsOut As Worksheet, nRow As Long, nCol As Integer, _
                           sCellJson As String, oRemittance As clsRemittance, _
                           iInvoiceIndex As Integer)
    Dim sCellType As String: sCellType = modUtils.JsonGetString(sCellJson, "cellType", CELLTYPE_LABEL)
    Dim sFieldMap As String: sFieldMap = modUtils.JsonGetString(sCellJson, "fieldMap", "")
    Dim sDefault As String: sDefault = modUtils.JsonGetString(sCellJson, "defaultValue", "")
    Dim sFormula As String: sFormula = modUtils.JsonGetString(sCellJson, "formula", "")

    Dim oCell As Range
    Set oCell = wsOut.Cells(nRow, nCol)

    Select Case UCase(sCellType)
        Case CELLTYPE_LABEL
            oCell.Value = sDefault

        Case CELLTYPE_BLANK
            oCell.Value = ""

        Case CELLTYPE_DATA
            ' Fill from remittance data
            oCell.Value = GetRemittanceField(sFieldMap, oRemittance, iInvoiceIndex, sDefault)

        Case CELLTYPE_LOOKUP
            ' Placeholder — will be filled by RunDataLookups
            oCell.Value = ""
            oCell.Interior.Color = COLOR_LOOKUP  ' Temporary marker

        Case CELLTYPE_CALC
            ' Formula cell — adjust row references for dynamic rows
            If sFormula <> "" Then
                Dim sAdjusted As String
                sAdjusted = AdjustFormulaRowRefs(sFormula, nRow)
                On Error Resume Next
                oCell.Formula = sAdjusted
                If Err.Number <> 0 Then
                    oCell.Value = sDefault
                    Err.Clear
                End If
                On Error GoTo 0
            ElseIf sDefault <> "" Then
                oCell.Value = sDefault
            End If
    End Select
End Sub

' ===========================================================
'  Map a fieldMap key to actual remittance data value
' ===========================================================
Private Function GetRemittanceField(sField As String, oRemittance As clsRemittance, _
                                      iIdx As Integer, sDefault As String) As Variant
    Select Case LCase(sField)
        Case "invoice_number", "invoicenumber", "inv_no"
            If iIdx >= 0 Then
                GetRemittanceField = oRemittance.InvoiceNumber(iIdx)
            Else
                GetRemittanceField = sDefault
            End If

        Case "invoice_amount", "invoiceamount", "inv_amount", "amount"
            If iIdx >= 0 Then
                GetRemittanceField = oRemittance.InvoiceAmount(iIdx)
            Else
                GetRemittanceField = oRemittance.TotalExtracted
            End If

        Case "invoice_date", "invoicedate", "inv_date"
            If iIdx >= 0 Then
                GetRemittanceField = oRemittance.InvoiceDate(iIdx)
            Else
                GetRemittanceField = sDefault
            End If

        Case "payment_date", "paymentdate", "remit_date"
            GetRemittanceField = oRemittance.PaymentDate

        Case "payment_amount", "paymentamount", "total_amount"
            GetRemittanceField = oRemittance.PaymentAmount

        Case "check_number", "checknumber", "eft_number", "reference"
            GetRemittanceField = oRemittance.CheckNumber

        Case "payor_name", "payorname"
            GetRemittanceField = oRemittance.PayorName

        Case "remittance_id", "remittanceid"
            GetRemittanceField = oRemittance.RemittanceId

        Case "total_extracted"
            GetRemittanceField = oRemittance.TotalExtracted

        Case "variance"
            GetRemittanceField = oRemittance.Variance

        Case "invoice_count"
            GetRemittanceField = oRemittance.InvoiceCount

        Case "row_number", "line_number"
            GetRemittanceField = iIdx + 1

        Case Else
            GetRemittanceField = sDefault
    End Select
End Function

' ===========================================================
'  Adjust formula row references for dynamic row expansion
'  Replaces placeholder row markers like {ROW} with actual row
' ===========================================================
Private Function AdjustFormulaRowRefs(sFormula As String, nActualRow As Long) As String
    ' Replace {ROW} placeholder with actual row number
    sFormula = Replace(sFormula, "{ROW}", CStr(nActualRow))
    AdjustFormulaRowRefs = sFormula
End Function

' ===========================================================
'  APPLY TEMPLATE FORMATTING to all cells
' ===========================================================
Private Sub ApplyTemplateFormatting(wsOut As Worksheet, oTemplate As clsTemplate, _
                                      nBodyStart As Long, nBodyEnd As Long, _
                                      nFooterStart As Long, nInvoices As Integer)
    Dim i As Integer
    For i = 0 To oTemplate.CellCount - 1
        Dim sCellJson As String: sCellJson = oTemplate.CellJson(i)
        Dim sSection As String: sSection = modUtils.JsonGetString(sCellJson, "section", "")
        Dim nRowOffset As Long: nRowOffset = modUtils.JsonGetNumber(sCellJson, "rowOffset", 1)
        Dim nCol As Integer: nCol = CInt(modUtils.JsonGetNumber(sCellJson, "col", 1))
        Dim sFormatJson As String: sFormatJson = modUtils.JsonGetRawValue(sCellJson, "formatJson", "{}")
        Dim bDynamic As Boolean: bDynamic = modUtils.JsonGetBool(sCellJson, "isDynamic", False)

        Dim oFmt As New clsCellFormat
        If sFormatJson <> "{}" And sFormatJson <> "" Then
            oFmt.FromJson sFormatJson
        Else
            ' Use section default
            Select Case sSection
                Case SECTION_HEADER: oFmt.FromJson oTemplate.DefaultHeaderFormatJson
                Case SECTION_BODY: oFmt.FromJson oTemplate.DefaultBodyFormatJson
                Case SECTION_FOOTER: oFmt.FromJson oTemplate.DefaultFooterFormatJson
            End Select
        End If

        If bDynamic Then
            ' Apply to all invoice body rows
            Dim iInv As Integer
            For iInv = 0 To nInvoices - 1
                Dim nAbsRow As Long: nAbsRow = nBodyStart + iInv * oTemplate.BodyRowCount + nRowOffset - 1
                oFmt.ApplyToRange wsOut.Cells(nAbsRow, nCol)
            Next iInv
        Else
            Dim nTargetRow As Long
            Select Case sSection
                Case SECTION_HEADER: nTargetRow = nRowOffset
                Case SECTION_FOOTER: nTargetRow = nFooterStart + nRowOffset - 1
                Case Else: nTargetRow = nBodyStart + nRowOffset - 1
            End Select
            oFmt.ApplyToRange wsOut.Cells(nTargetRow, nCol)
        End If
    Next i

    ' Set column widths
    Dim c As Integer
    For c = 1 To oTemplate.ColumnCount
        Dim dWidth As Single: dWidth = oTemplate.GetColumnWidth(c)
        If dWidth = 0 Then
            wsOut.Columns(c).AutoFit
        ElseIf dWidth > 0 Then
            wsOut.Columns(c).ColumnWidth = dWidth
        End If
    Next c
End Sub

' ===========================================================
'  RUN DATA LOOKUPS via Power Query / Data Model
' ===========================================================
Private Sub RunDataLookups(oTemplate As clsTemplate, oRemittance As clsRemittance, _
                             oPayor As clsPayor, wsOut As Worksheet)
    Dim oConfig As Object
    Set oConfig = modUtils.SimpleJsonToDictionary(oPayor.LookupConfigJson)

    Dim iInv As Integer
    For iInv = 0 To oRemittance.InvoiceCount - 1
        Dim sInvNo As String: sInvNo = oRemittance.InvoiceNumber(iInv)
        Dim oLookupResult As Object
        Set oLookupResult = modDataLookup.LookupInvoice(sInvNo, oConfig)

        If oLookupResult Is Nothing Then
            oRemittance.AddWarning "No lookup result for invoice: " & sInvNo
            GoTo NextLookup
        End If

        ' Find lookup cells for this invoice in the template
        Dim nBaseRow As Long
        nBaseRow = oTemplate.BodyStartRow + (iInv * oTemplate.BodyRowCount)

        Dim c As Integer
        For c = 0 To oTemplate.CellCount - 1
            Dim sCellJson As String: sCellJson = oTemplate.CellJson(c)
            If modUtils.JsonGetString(sCellJson, "cellType", "") = CELLTYPE_LOOKUP Then
                Dim sField As String: sField = modUtils.JsonGetString(sCellJson, "fieldMap", "")
                If oLookupResult.Exists(sField) Then
                    Dim nRow As Long
                    If modUtils.JsonGetBool(sCellJson, "isDynamic", False) Then
                        nRow = nBaseRow + modUtils.JsonGetNumber(sCellJson, "rowOffset", 1) - 1
                    Else
                        nRow = modUtils.JsonGetNumber(sCellJson, "rowOffset", 1)
                    End If
                    Dim nCol As Integer: nCol = CInt(modUtils.JsonGetNumber(sCellJson, "col", 1))
                    wsOut.Cells(nRow, nCol).Value = oLookupResult(sField)
                    ' Clear the placeholder lookup color
                    wsOut.Cells(nRow, nCol).Interior.Pattern = xlNone
                End If
            End If
        Next c
NextLookup:
    Next iInv

    oRemittance.AddAuditLine "Data lookups complete for " & oRemittance.InvoiceCount & " invoice(s)."
End Sub

' ===========================================================
'  APPLY PRINT SETTINGS
' ===========================================================
Private Sub ApplyPrintSettings(wsOut As Worksheet, oPayor As clsPayor, _
                                 oTemplate As clsTemplate)
    On Error Resume Next
    Dim oPrint As Object
    Set oPrint = modUtils.SimpleJsonToDictionary(oPayor.PrintSettingsJson)

    With wsOut.PageSetup
        ' Orientation
        Dim sOrient As String
        sOrient = LCase(modUtils.JsonGetString(oPayor.PrintSettingsJson, "orientation", "portrait"))
        .Orientation = IIf(sOrient = "landscape", xlLandscape, xlPortrait)

        ' Margins (in inches)
        .TopMargin = Application.InchesToPoints( _
            CDbl(modUtils.JsonGetNumber(oPayor.PrintSettingsJson, "marginTop", 0.75)))
        .BottomMargin = Application.InchesToPoints( _
            CDbl(modUtils.JsonGetNumber(oPayor.PrintSettingsJson, "marginBottom", 0.75)))
        .LeftMargin = Application.InchesToPoints( _
            CDbl(modUtils.JsonGetNumber(oPayor.PrintSettingsJson, "marginLeft", 0.7)))
        .RightMargin = Application.InchesToPoints( _
            CDbl(modUtils.JsonGetNumber(oPayor.PrintSettingsJson, "marginRight", 0.7)))

        ' Fit to page
        If modUtils.JsonGetBool(oPayor.PrintSettingsJson, "fitToPage", True) Then
            .FitToPagesWide = CInt(modUtils.JsonGetNumber(oPayor.PrintSettingsJson, "fitToWidth", 1))
            .FitToPagesTall = CInt(modUtils.JsonGetNumber(oPayor.PrintSettingsJson, "fitToHeight", 0))
            .Zoom = False
        End If

        ' Center on page
        .CenterHorizontally = modUtils.JsonGetBool(oPayor.PrintSettingsJson, "centerHorizontal", True)
        .CenterVertically = modUtils.JsonGetBool(oPayor.PrintSettingsJson, "centerVertical", False)

        ' Header/footer text
        Dim sHdr As String: sHdr = modUtils.JsonGetString(oPayor.PrintSettingsJson, "headerText", "")
        Dim sFtr As String: sFtr = modUtils.JsonGetString(oPayor.PrintSettingsJson, "footerText", "Page &P of &N")
        If sHdr <> "" Then .CenterHeader = sHdr
        If sFtr <> "" Then .CenterFooter = sFtr

        ' Print area
        Dim nLastRow As Long: nLastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
        Dim nLastCol As Integer: nLastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
        .PrintArea = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(nLastRow, nLastCol)).Address
    End With

    ' Freeze panes
    Dim nFreezeRow As Integer: nFreezeRow = oTemplate.FreezePaneRow
    Dim nFreezeCol As Integer: nFreezeCol = oTemplate.FreezePaneCol
    If nFreezeRow > 0 Or nFreezeCol > 0 Then
        wsOut.Activate
        Application.Goto wsOut.Cells(1, 1), True
        wsOut.Cells(nFreezeRow + 1, IIf(nFreezeCol > 0, nFreezeCol + 1, 1)).Select
        ActiveWindow.FreezePanes = True
    End If

    On Error GoTo 0
End Sub

' ===========================================================
'  FLAG lookup errors with red fill
' ===========================================================
Private Sub FlagLookupErrors(wsOut As Worksheet, oRemittance As clsRemittance)
    ' Find cells still colored with lookup placeholder color
    On Error Resume Next
    Dim rng As Range
    For Each rng In wsOut.UsedRange
        If rng.Interior.Color = COLOR_LOOKUP And rng.Value = "" Then
            rng.Interior.Color = COLOR_ERROR       ' Red flag
            rng.Value = "#LOOKUP_FAILED"
        End If
    Next rng
    On Error GoTo 0
End Sub

' ===========================================================
'  FLAG duplicate invoices in output
' ===========================================================
Private Sub FlagDuplicates(wsOut As Worksheet, oRemittance As clsRemittance, _
                              oTemplate As clsTemplate)
    If Not oRemittance.HasDuplicates Then Exit Sub

    Dim iInv As Integer
    For iInv = 0 To oRemittance.InvoiceCount - 1
        If oRemittance.IsDuplicate(iInv) Then
            Dim nRow As Long
            nRow = oTemplate.BodyStartRow + (iInv * oTemplate.BodyRowCount)
            Dim nLastCol As Integer: nLastCol = oTemplate.ColumnCount
            ' Orange fill for duplicate rows
            wsOut.Range( _
                wsOut.Cells(nRow, 1), _
                wsOut.Cells(nRow + oTemplate.BodyRowCount - 1, nLastCol) _
            ).Interior.Color = RGB(255, 165, 0)  ' Orange
        End If
    Next iInv
End Sub

' ===========================================================
'  Create clean output sheet
' ===========================================================
Private Function CreateOutputSheet(sPayorName As String) As Worksheet
    Dim sSheetName As String
    ' Build a unique name: PayorName + date
    sSheetName = Left(sPayorName, 20) & "_" & Format(Now(), "mmddyy_hhmmss")
    ' Sanitize for sheet name
    sSheetName = Replace(sSheetName, "/", "-")
    sSheetName = Replace(sSheetName, "\", "-")
    sSheetName = Replace(sSheetName, ":", "-")
    sSheetName = Replace(sSheetName, "[", "")
    sSheetName = Replace(sSheetName, "]", "")
    sSheetName = Replace(sSheetName, "*", "")
    sSheetName = Replace(sSheetName, "?", "")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = Left(sSheetName, 31)  ' Max 31 chars
    Set CreateOutputSheet = ws
End Function
