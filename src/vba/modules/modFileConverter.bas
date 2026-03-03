Attribute VB_Name = "modFileConverter"
Option Explicit

' ============================================================
'  modFileConverter — Input File Handling & Conversion Layer
'
'  Accepts remittance files in any of these formats:
'    .xlsx  — read directly
'    .xls (genuine OLE2) — convert via Excel
'    .xls (fake HTML) — detect via magic bytes, parse HTML
'    .csv  — parse with auto-delimiter detection
'
'  Returns: path to a normalized .xlsx file in TEMP folder.
'  The original file is NEVER modified.
' ============================================================

' ===========================================================
'  PUBLIC ENTRY POINT
'  Returns normalized .xlsx path, or "" on failure.
'  Caller must delete the temp file after use.
' ===========================================================
Public Function ConvertToXlsx(sSourcePath As String, _
                               oRemittance As clsRemittance) As String
    If Not modUtils.FileExists(sSourcePath) Then
        modUtils.ShowError "File not found: " & sSourcePath
        ConvertToXlsx = ""
        Exit Function
    End If

    Dim sExt As String
    sExt = LCase(Right(sSourcePath, Len(sSourcePath) - InStrRev(sSourcePath, ".")))

    Dim sDetectedType As String
    Dim sResult As String

    Select Case sExt
        Case "xlsx", "xlsm"
            sDetectedType = FILETYPE_XLSX
            sResult = CopyToTemp(sSourcePath, "xlsx")
            oRemittance.AddAuditLine "Input: .xlsx — copied to temp, no conversion needed."

        Case "xls"
            ' Must detect: genuine OLE2 binary, or fake HTML
            If IsHtmlXls(sSourcePath) Then
                sDetectedType = FILETYPE_XLS_HTML
                oRemittance.AddAuditLine "Detected: .xls is actually HTML. Converting via HTML parser."
                sResult = ConvertHtmlXlsToXlsx(sSourcePath, oRemittance)
            Else
                sDetectedType = FILETYPE_XLS_GENUINE
                oRemittance.AddAuditLine "Detected: genuine OLE2 .xls binary. Converting via Excel."
                sResult = ConvertGenuineXlsToXlsx(sSourcePath, oRemittance)
            End If

        Case "csv", "txt"
            sDetectedType = FILETYPE_CSV
            oRemittance.AddAuditLine "Input: .csv — converting with delimiter detection."
            sResult = ConvertCsvToXlsx(sSourcePath, oRemittance)

        Case Else
            ' Unknown — attempt as xlsx
            sDetectedType = FILETYPE_XLSX
            sResult = CopyToTemp(sSourcePath, "xlsx")
            oRemittance.AddAuditLine "Warning: Unknown extension ." & sExt & " — treating as xlsx."
    End Select

    oRemittance.SourceFileType = sDetectedType

    If sResult = "" Then
        modUtils.WriteAuditLog "ERROR", "FileConverter", _
            "Conversion failed for: " & sSourcePath, , sSourcePath, "FAIL"
    Else
        modUtils.WriteAuditLog "CONVERT", "FileConverter", _
            "Converted " & sDetectedType & " → " & sResult, , sSourcePath, "OK"
    End If

    ConvertToXlsx = sResult
End Function

' ===========================================================
'  Detect if a .xls file is actually an HTML document
'  Real OLE2 files start with D0 CF 11 E0 (magic bytes)
'  HTML files start with <html, <!DOCTYPE, <table, etc.
' ===========================================================
Public Function IsHtmlXls(sFilePath As String) As Boolean
    IsHtmlXls = False
    On Error GoTo DetectErr

    Dim fileNum As Integer
    fileNum = FreeFile()
    Open sFilePath For Binary Access Read As #fileNum

    Dim header As String * 16
    Get #fileNum, , header
    Close #fileNum

    ' Check for OLE2 magic: D0 CF 11 E0 A1 B1 1A E1
    If Left(header, 4) = OLE2_MAGIC Then
        IsHtmlXls = False
        Exit Function
    End If

    ' Check for HTML markers
    Dim hLower As String
    hLower = LCase(Trim(Left(header, 16)))
    If InStr(hLower, "<html") > 0 Or _
       InStr(hLower, "<!doc") > 0 Or _
       InStr(hLower, "<tabl") > 0 Or _
       InStr(hLower, "<?xml") > 0 Or _
       Left(hLower, 3) = Chr(239) & Chr(187) & Chr(191) Then  ' BOM
        IsHtmlXls = True
        Exit Function
    End If

    ' Read more of the file to look for HTML structure
    Dim fullHeader As String * 512
    fileNum = FreeFile()
    Open sFilePath For Binary Access Read As #fileNum
    Get #fileNum, , fullHeader
    Close #fileNum

    Dim fhLower As String
    fhLower = LCase(fullHeader)
    If InStr(fhLower, "<html") > 0 Or _
       InStr(fhLower, "<table") > 0 Or _
       InStr(fhLower, "<!doctype html") > 0 Then
        IsHtmlXls = True
    End If

    Exit Function
DetectErr:
    IsHtmlXls = False
    Err.Clear
End Function

' ===========================================================
'  Convert genuine .xls to .xlsx via Excel open/save-as
' ===========================================================
Private Function ConvertGenuineXlsToXlsx(sSource As String, _
                                           oRemittance As clsRemittance) As String
    Dim sOut As String
    sOut = modUtils.GetTempFilePath("xlsx")

    On Error GoTo ConvErr

    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Open(sSource, _
                ReadOnly:=True, _
                IgnoreReadOnlyRecommended:=True, _
                Notify:=False, _
                CorruptLoad:=xlRepairFile)

    wb.SaveAs Filename:=sOut, _
              FileFormat:=xlOpenXMLWorkbook, _
              CreateBackup:=False
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    Set wb = Nothing
    oRemittance.AddAuditLine "Converted genuine .xls → " & sOut
    ConvertGenuineXlsToXlsx = sOut
    Exit Function

ConvErr:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    oRemittance.AddAuditLine "ERROR converting .xls: " & Err.Description
    oRemittance.AddWarning "Failed to convert .xls file: " & Err.Description
    ConvertGenuineXlsToXlsx = ""
    Err.Clear
End Function

' ===========================================================
'  Convert fake-HTML .xls to .xlsx by parsing HTML tables
' ===========================================================
Private Function ConvertHtmlXlsToXlsx(sSource As String, _
                                        oRemittance As clsRemittance) As String
    Dim sOut As String
    sOut = modUtils.GetTempFilePath("xlsx")

    On Error GoTo HtmlErr

    ' Read entire file as text
    Dim sHtml As String
    sHtml = ReadFileAsText(sSource)
    If sHtml = "" Then
        oRemittance.AddWarning "Could not read HTML content from: " & sSource
        ConvertHtmlXlsToXlsx = ""
        Exit Function
    End If

    ' Parse HTML tables
    Dim tables() As String
    tables = ExtractHtmlTables(sHtml)
    If UBound(tables) < 0 Or (UBound(tables) = 0 And tables(0) = "") Then
        oRemittance.AddWarning "No HTML table found in file: " & sSource
        ConvertHtmlXlsToXlsx = ""
        Exit Function
    End If

    ' Create new workbook and write table data
    Dim wbOut As Workbook
    Application.DisplayAlerts = False
    Set wbOut = Application.Workbooks.Add

    ' Process first table (or all tables as separate sheets)
    Dim t As Integer
    For t = 0 To UBound(tables)
        If Trim(tables(t)) <> "" Then
            Dim ws As Worksheet
            If t = 0 Then
                Set ws = wbOut.Sheets(1)
                ws.Name = "Sheet1"
            Else
                Set ws = wbOut.Sheets.Add(After:=wbOut.Sheets(wbOut.Sheets.Count))
                ws.Name = "Sheet" & CStr(t + 1)
            End If
            WriteHtmlTableToSheet tables(t), ws, oRemittance
        End If
    Next t

    wbOut.SaveAs Filename:=sOut, _
                 FileFormat:=xlOpenXMLWorkbook, _
                 CreateBackup:=False
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Set wbOut = Nothing

    oRemittance.AddAuditLine "HTML-as-XLS converted: " & UBound(tables) + 1 & " table(s) → " & sOut
    ConvertHtmlXlsToXlsx = sOut
    Exit Function

HtmlErr:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    On Error GoTo 0
    oRemittance.AddAuditLine "ERROR parsing HTML-XLS at: " & Err.Description
    oRemittance.AddWarning "HTML conversion failed: " & Err.Description
    ConvertHtmlXlsToXlsx = ""
    Err.Clear
End Function

' ===========================================================
'  Convert .csv to .xlsx with delimiter auto-detection
' ===========================================================
Private Function ConvertCsvToXlsx(sSource As String, _
                                    oRemittance As clsRemittance) As String
    Dim sOut As String
    sOut = modUtils.GetTempFilePath("xlsx")

    On Error GoTo CsvErr

    ' Read file lines
    Dim lines() As String
    lines = ReadFileLines(sSource)
    If UBound(lines) < 0 Then
        oRemittance.AddWarning "Empty or unreadable CSV file: " & sSource
        ConvertCsvToXlsx = ""
        Exit Function
    End If

    ' Detect delimiter from first non-empty line
    Dim sDelim As String: sDelim = DELIM_COMMA
    Dim i As Integer
    For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            sDelim = modUtils.DetectDelimiter(lines(i))
            Exit For
        End If
    Next i
    oRemittance.AddAuditLine "CSV delimiter detected: [" & ShowDelim(sDelim) & "]"

    ' Create workbook
    Dim wbOut As Workbook
    Application.DisplayAlerts = False
    Set wbOut = Application.Workbooks.Add
    Dim ws As Worksheet
    Set ws = wbOut.Sheets(1)
    ws.Name = "Data"

    ' Write each CSV row to sheet
    Dim nRow As Long: nRow = 1
    Dim nMaxCol As Integer: nMaxCol = 0

    For i = 0 To UBound(lines)
        Dim sLine As String: sLine = lines(i)
        If Trim(sLine) = "" Then GoTo NextLine

        Dim fields() As String
        fields = modUtils.SplitDelimited(sLine, sDelim)

        Dim nCol As Integer
        For nCol = 0 To UBound(fields)
            Dim sField As String: sField = Trim(fields(nCol))
            ' Try to parse as number
            Dim numVal As Double
            On Error Resume Next
            numVal = CDbl(Replace(sField, ",", ""))
            If Err.Number = 0 And IsNumeric(Replace(sField, ",", "")) Then
                ws.Cells(nRow, nCol + 1).Value = numVal
            Else
                ws.Cells(nRow, nCol + 1).Value = sField
            End If
            On Error GoTo CsvErr
            If nCol + 1 > nMaxCol Then nMaxCol = nCol + 1
        Next nCol

        nRow = nRow + 1
NextLine:
    Next i

    oRemittance.AddAuditLine "CSV rows written: " & (nRow - 1) & ", columns: " & nMaxCol

    wbOut.SaveAs Filename:=sOut, _
                 FileFormat:=xlOpenXMLWorkbook, _
                 CreateBackup:=False
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Set wbOut = Nothing

    ConvertCsvToXlsx = sOut
    Exit Function

CsvErr:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    On Error GoTo 0
    oRemittance.AddAuditLine "ERROR converting CSV: " & Err.Description
    oRemittance.AddWarning "CSV conversion failed: " & Err.Description
    ConvertCsvToXlsx = ""
    Err.Clear
End Function

' ===========================================================
'  Copy a file to temp (no conversion needed)
' ===========================================================
Private Function CopyToTemp(sSource As String, sExt As String) As String
    Dim sDest As String
    sDest = modUtils.GetTempFilePath(sExt)
    On Error Resume Next
    FileCopy sSource, sDest
    If Err.Number <> 0 Then
        CopyToTemp = ""
        Err.Clear
    Else
        CopyToTemp = sDest
    End If
    On Error GoTo 0
End Function

' ===========================================================
'  Read entire text file as single string
' ===========================================================
Private Function ReadFileAsText(sPath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile()
    Dim result As String

    On Error GoTo ReadErr
    Open sPath For Input As #fileNum
    Dim sLine As String
    Dim allLines As String
    allLines = ""
    Do While Not EOF(fileNum)
        Line Input #fileNum, sLine
        allLines = allLines & sLine & vbLf
    Loop
    Close #fileNum
    ReadFileAsText = allLines
    Exit Function
ReadErr:
    On Error Resume Next
    Close #fileNum
    ReadFileAsText = ""
    Err.Clear
End Function

' ===========================================================
'  Read text file into array of lines
' ===========================================================
Private Function ReadFileLines(sPath As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim count As Integer: count = 0

    Dim fileNum As Integer
    fileNum = FreeFile()
    On Error GoTo LinesErr
    Open sPath For Input As #fileNum
    Dim sLine As String
    Do While Not EOF(fileNum)
        Line Input #fileNum, sLine
        ReDim Preserve result(count)
        result(count) = sLine
        count = count + 1
    Loop
    Close #fileNum
    ReadFileLines = result
    Exit Function

LinesErr:
    On Error Resume Next
    Close #fileNum
    ReDim result(0)
    ReadFileLines = result
    Err.Clear
End Function

' ===========================================================
'  Extract HTML table contents from HTML string
'  Returns array of table inner HTML strings
' ===========================================================
Private Function ExtractHtmlTables(sHtml As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim count As Integer: count = 0

    Dim sLower As String: sLower = LCase(sHtml)
    Dim nStart As Long: nStart = 1
    Dim nTableStart As Long
    Dim nTableEnd As Long

    Do
        ' Find <table
        nTableStart = InStr(nStart, sLower, "<table")
        If nTableStart = 0 Then Exit Do

        ' Find </table>
        nTableEnd = InStr(nTableStart, sLower, "</table>")
        If nTableEnd = 0 Then Exit Do

        nTableEnd = nTableEnd + Len("</table>") - 1
        Dim sTable As String
        sTable = Mid(sHtml, nTableStart, nTableEnd - nTableStart + 1)

        ReDim Preserve result(count)
        result(count) = sTable
        count = count + 1

        nStart = nTableEnd + 1
    Loop

    ExtractHtmlTables = result
End Function

' ===========================================================
'  Parse an HTML table string and write cell values to a sheet
' ===========================================================
Private Sub WriteHtmlTableToSheet(sTableHtml As String, ws As Worksheet, _
                                    oRemittance As clsRemittance)
    Dim sLower As String: sLower = LCase(sTableHtml)
    Dim nRow As Long: nRow = 1
    Dim nPos As Long: nPos = 1
    Dim nTrStart As Long, nTrEnd As Long
    Dim errorCount As Integer: errorCount = 0

    Do
        ' Find <tr
        nTrStart = InStr(nPos, sLower, "<tr")
        If nTrStart = 0 Then Exit Do

        ' Find </tr>
        nTrEnd = InStr(nTrStart, sLower, "</tr>")
        If nTrEnd = 0 Then Exit Do

        Dim sTrContent As String
        sTrContent = Mid(sTableHtml, nTrStart, nTrEnd - nTrStart + Len("</tr>"))

        ' Extract cells (<td> and <th>)
        Dim nCol As Integer: nCol = 1
        Dim sCellLower As String: sCellLower = LCase(sTrContent)
        Dim nCellPos As Long: nCellPos = 1
        Dim nCellStart As Long, nCellEnd As Long

        Do
            ' Find <td or <th
            Dim nTd As Long: nTd = InStr(nCellPos, sCellLower, "<td")
            Dim nTh As Long: nTh = InStr(nCellPos, sCellLower, "<th")
            If nTd = 0 And nTh = 0 Then Exit Do

            Dim nCellTagStart As Long
            If nTd = 0 Then
                nCellTagStart = nTh
            ElseIf nTh = 0 Then
                nCellTagStart = nTd
            Else
                nCellTagStart = IIf(nTd < nTh, nTd, nTh)
            End If

            ' Find closing > of opening tag
            Dim nTagClose As Long
            nTagClose = InStr(nCellTagStart, sCellLower, ">")
            If nTagClose = 0 Then Exit Do

            ' Find closing </td> or </th>
            Dim nTdClose As Long: nTdClose = InStr(nTagClose, sCellLower, "</td>")
            Dim nThClose As Long: nThClose = InStr(nTagClose, sCellLower, "</th>")
            Dim nEndTag As Long
            If nTdClose = 0 And nThClose = 0 Then
                nCellPos = nTagClose + 1
                GoTo NextCell
            ElseIf nTdClose = 0 Then
                nEndTag = nThClose
            ElseIf nThClose = 0 Then
                nEndTag = nTdClose
            Else
                nEndTag = IIf(nTdClose < nThClose, nTdClose, nThClose)
            End If

            ' Extract inner HTML
            Dim sCellInner As String
            sCellInner = Mid(sTrContent, nTagClose + 1, nEndTag - nTagClose - 1)

            ' Strip HTML tags
            Dim sCellText As String
            sCellText = StripHtmlTags(sCellInner)
            sCellText = DecodeHtmlEntities(sCellText)
            sCellText = Trim(sCellText)

            On Error Resume Next
            ws.Cells(nRow, nCol).Value = sCellText
            On Error GoTo 0

            nCol = nCol + 1
            nCellPos = nEndTag + 5  ' Move past </td> or </th>
NextCell:
        Loop

        nRow = nRow + 1
        nPos = nTrEnd + 5  ' Move past </tr>
    Loop

    oRemittance.AddAuditLine "HTML table parsed: " & (nRow - 1) & " rows written to " & ws.Name
End Sub

' ===========================================================
'  Strip all HTML tags from a string
' ===========================================================
Private Function StripHtmlTags(sHtml As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "<[^>]+>"
    re.Global = True
    StripHtmlTags = re.Replace(sHtml, "")
    Set re = Nothing
End Function

' ===========================================================
'  Decode common HTML entities
' ===========================================================
Private Function DecodeHtmlEntities(s As String) As String
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&apos;", "'")
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&#160;", " ")
    s = Replace(s, "&#38;", "&")
    DecodeHtmlEntities = s
End Function

' ===========================================================
'  Display-friendly delimiter name
' ===========================================================
Private Function ShowDelim(s As String) As String
    Select Case s
        Case DELIM_COMMA: ShowDelim = "comma"
        Case DELIM_TAB: ShowDelim = "tab"
        Case DELIM_PIPE: ShowDelim = "pipe"
        Case DELIM_SEMICOLON: ShowDelim = "semicolon"
        Case DELIM_SPACE: ShowDelim = "space"
        Case Else: ShowDelim = s
    End Select
End Function
