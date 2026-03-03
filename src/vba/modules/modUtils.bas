Attribute VB_Name = "modUtils"
Option Explicit

' ============================================================
'  modUtils — Utility Functions
'  String helpers, JSON parser, regex, file I/O, audit log
' ============================================================

' ===========================================================
'  AUDIT LOG
' ===========================================================

Public Sub WriteAuditLog(sAction As String, sContext As String, sDetail As String, _
                          Optional sPayor As String = "", _
                          Optional sFile As String = "", _
                          Optional sStatus As String = "OK")
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_AUDIT_LOG)
    If ws Is Nothing Then Exit Sub

    ' Initialize header if first row is empty
    If ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_TIMESTAMP).Value = "" Then
        ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_TIMESTAMP).Value = "Timestamp"
        ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_ACTION).Value = "Action"
        ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_PAYOR).Value = "Payor"
        ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_FILE).Value = "File"
        ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_DETAIL).Value = "Detail"
        ws.Cells(AUDIT_HEADER_ROW, AUDIT_COL_STATUS).Value = "Status"
    End If

    ' Find next empty row
    Dim nRow As Long
    nRow = ws.Cells(ws.Rows.Count, AUDIT_COL_TIMESTAMP).End(xlUp).Row + 1
    If nRow <= AUDIT_HEADER_ROW Then nRow = AUDIT_DATA_START_ROW

    ' Trim old entries if over limit
    If nRow > AUDIT_MAX_ROWS + AUDIT_DATA_START_ROW Then
        ws.Rows(AUDIT_DATA_START_ROW).Delete
        nRow = nRow - 1
    End If

    ws.Cells(nRow, AUDIT_COL_TIMESTAMP).Value = Now()
    ws.Cells(nRow, AUDIT_COL_TIMESTAMP).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nRow, AUDIT_COL_ACTION).Value = sAction
    ws.Cells(nRow, AUDIT_COL_PAYOR).Value = sPayor
    ws.Cells(nRow, AUDIT_COL_FILE).Value = sFile
    ws.Cells(nRow, AUDIT_COL_DETAIL).Value = sContext & ": " & sDetail
    ws.Cells(nRow, AUDIT_COL_STATUS).Value = sStatus
    On Error GoTo 0
End Sub

' ===========================================================
'  SIMPLE JSON PARSER HELPERS
'  Handles flat JSON objects (not nested arrays as values
'  except via JsonGetRawValue).
' ===========================================================

' Get a string value from a flat JSON object
Public Function JsonGetString(sJson As String, sKey As String, sDefault As String) As String
    Dim sVal As String
    sVal = JsonExtractValue(sJson, sKey)
    If sVal = vbNullString Then
        JsonGetString = sDefault
    Else
        ' Strip surrounding quotes if present
        If Left(sVal, 1) = """" And Right(sVal, 1) = """" Then
            sVal = Mid(sVal, 2, Len(sVal) - 2)
        End If
        ' Unescape common sequences
        sVal = Replace(sVal, "\""", """")
        sVal = Replace(sVal, "\\", "\")
        sVal = Replace(sVal, "\n", vbLf)
        sVal = Replace(sVal, "\r", vbCr)
        JsonGetString = sVal
    End If
End Function

' Get a numeric value from a flat JSON object
Public Function JsonGetNumber(sJson As String, sKey As String, dDefault As Double) As Double
    Dim sVal As String
    sVal = JsonExtractValue(sJson, sKey)
    If sVal = vbNullString Or Trim(sVal) = "null" Then
        JsonGetNumber = dDefault
    Else
        On Error Resume Next
        JsonGetNumber = CDbl(Trim(sVal))
        If Err.Number <> 0 Then JsonGetNumber = dDefault
        On Error GoTo 0
    End If
End Function

' Get a boolean value from a flat JSON object
Public Function JsonGetBool(sJson As String, sKey As String, bDefault As Boolean) As Boolean
    Dim sVal As String
    sVal = Trim(JsonExtractValue(sJson, sKey))
    Select Case LCase(sVal)
        Case "true", "1": JsonGetBool = True
        Case "false", "0": JsonGetBool = False
        Case Else: JsonGetBool = bDefault
    End Select
End Function

' Get a raw value (could be nested object/array) from JSON
Public Function JsonGetRawValue(sJson As String, sKey As String, sDefault As String) As String
    Dim sVal As String
    sVal = JsonExtractRawValue(sJson, sKey)
    If sVal = vbNullString Then
        JsonGetRawValue = sDefault
    Else
        JsonGetRawValue = sVal
    End If
End Function

' Core extractor — finds "key": value and returns the raw value string
Private Function JsonExtractValue(sJson As String, sKey As String) As String
    Dim sSearch As String
    sSearch = """" & sKey & """:"
    Dim nPos As Long
    nPos = InStr(sJson, sSearch)
    If nPos = 0 Then
        JsonExtractValue = vbNullString
        Exit Function
    End If

    nPos = nPos + Len(sSearch)
    ' Skip whitespace
    Do While nPos <= Len(sJson) And Mid(sJson, nPos, 1) = " "
        nPos = nPos + 1
    Loop

    Dim ch As String: ch = Mid(sJson, nPos, 1)

    If ch = """" Then
        ' String value — find closing quote (handle escapes)
        Dim nEnd As Long: nEnd = nPos + 1
        Do While nEnd <= Len(sJson)
            If Mid(sJson, nEnd, 1) = "\" Then
                nEnd = nEnd + 2  ' Skip escaped character
            ElseIf Mid(sJson, nEnd, 1) = """" Then
                Exit Do
            Else
                nEnd = nEnd + 1
            End If
        Loop
        JsonExtractValue = Mid(sJson, nPos, nEnd - nPos + 1)
    ElseIf ch = "{" Or ch = "[" Then
        ' Nested object/array — skip (use JsonExtractRawValue for these)
        JsonExtractValue = vbNullString
    Else
        ' Numeric / boolean / null — read until comma or }
        Dim nEnd2 As Long: nEnd2 = nPos
        Do While nEnd2 <= Len(sJson)
            Dim c As String: c = Mid(sJson, nEnd2, 1)
            If c = "," Or c = "}" Then Exit Do
            nEnd2 = nEnd2 + 1
        Loop
        JsonExtractValue = Trim(Mid(sJson, nPos, nEnd2 - nPos))
    End If
End Function

' Extractor for nested objects/arrays — returns the whole nested value
Private Function JsonExtractRawValue(sJson As String, sKey As String) As String
    Dim sSearch As String
    sSearch = """" & sKey & """:"
    Dim nPos As Long
    nPos = InStr(sJson, sSearch)
    If nPos = 0 Then
        JsonExtractRawValue = vbNullString
        Exit Function
    End If

    nPos = nPos + Len(sSearch)
    Do While nPos <= Len(sJson) And Mid(sJson, nPos, 1) = " "
        nPos = nPos + 1
    Loop

    Dim ch As String: ch = Mid(sJson, nPos, 1)
    If ch = "{" Or ch = "[" Then
        Dim closeChar As String
        closeChar = IIf(ch = "{", "}", "]")
        Dim depth As Integer: depth = 1
        Dim nEnd As Long: nEnd = nPos + 1
        Dim inString As Boolean: inString = False
        Do While nEnd <= Len(sJson) And depth > 0
            Dim c As String: c = Mid(sJson, nEnd, 1)
            If c = "\" And inString Then
                nEnd = nEnd + 2
            ElseIf c = """" Then
                inString = Not inString
                nEnd = nEnd + 1
            ElseIf Not inString Then
                If c = ch Or c = IIf(ch = "{", "[", "{") Then depth = depth + 1
                If c = closeChar Or c = IIf(closeChar = "}", "]", "}") Then depth = depth - 1
                nEnd = nEnd + 1
            Else
                nEnd = nEnd + 1
            End If
        Loop
        JsonExtractRawValue = Mid(sJson, nPos, nEnd - nPos - 1)
    Else
        ' Scalar — delegate to regular extractor
        JsonExtractRawValue = JsonExtractValue(sJson, sKey)
    End If
End Function

' ===========================================================
'  Split a JSON array string ["a","b",...] into individual items
'  Handles nested objects within the array.
' ===========================================================
Public Function SplitJsonArray(sArr As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim count As Integer: count = 0

    sArr = Trim(sArr)
    If Left(sArr, 1) = "[" Then sArr = Mid(sArr, 2)
    If Right(sArr, 1) = "]" Then sArr = Left(sArr, Len(sArr) - 1)
    sArr = Trim(sArr)
    If sArr = "" Then
        SplitJsonArray = result
        Exit Function
    End If

    Dim depth As Integer: depth = 0
    Dim inStr As Boolean: inStr = False
    Dim startPos As Long: startPos = 1
    Dim i As Long

    For i = 1 To Len(sArr)
        Dim c As String: c = Mid(sArr, i, 1)
        If c = "\" And inStr Then
            i = i + 1  ' Skip escaped char
        ElseIf c = """" Then
            inStr = Not inStr
        ElseIf Not inStr Then
            If c = "{" Or c = "[" Then
                depth = depth + 1
            ElseIf c = "}" Or c = "]" Then
                depth = depth - 1
            ElseIf c = "," And depth = 0 Then
                ReDim Preserve result(count)
                result(count) = Trim(Mid(sArr, startPos, i - startPos))
                count = count + 1
                startPos = i + 1
            End If
        End If
    Next i

    ' Last item
    If startPos <= Len(sArr) Then
        ReDim Preserve result(count)
        result(count) = Trim(Mid(sArr, startPos))
        count = count + 1
    End If

    SplitJsonArray = result
End Function

' ===========================================================
'  Convert a simple flat JSON object to a Scripting.Dictionary
' ===========================================================
Public Function SimpleJsonToDictionary(sJson As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    sJson = Trim(sJson)
    If Left(sJson, 1) = "{" Then sJson = Mid(sJson, 2)
    If Right(sJson, 1) = "}" Then sJson = Left(sJson, Len(sJson) - 1)

    ' Very simple key:value splitting — does not handle nested objects
    Dim inStr As Boolean: inStr = False
    Dim depth As Integer: depth = 0
    Dim startPos As Long: startPos = 1
    Dim pairs() As String
    ReDim pairs(0)
    Dim pc As Integer: pc = 0
    Dim i As Long

    For i = 1 To Len(sJson)
        Dim c As String: c = Mid(sJson, i, 1)
        If c = "\" And inStr Then
            i = i + 1
        ElseIf c = """" Then
            inStr = Not inStr
        ElseIf Not inStr Then
            If c = "{" Or c = "[" Then depth = depth + 1
            If c = "}" Or c = "]" Then depth = depth - 1
            If c = "," And depth = 0 Then
                ReDim Preserve pairs(pc)
                pairs(pc) = Trim(Mid(sJson, startPos, i - startPos))
                pc = pc + 1
                startPos = i + 1
            End If
        End If
    Next i
    If startPos <= Len(sJson) Then
        ReDim Preserve pairs(pc)
        pairs(pc) = Trim(Mid(sJson, startPos))
    End If

    Dim j As Integer
    For j = 0 To pc
        Dim pair As String: pair = Trim(pairs(j))
        Dim colonPos As Long: colonPos = InStr(pair, ":")
        If colonPos > 0 Then
            Dim sKey As String: sKey = Trim(Left(pair, colonPos - 1))
            sKey = Replace(sKey, """", "")
            Dim sVal As String: sVal = Trim(Mid(pair, colonPos + 1))
            If Left(sVal, 1) = """" And Right(sVal, 1) = """" Then
                sVal = Mid(sVal, 2, Len(sVal) - 2)
                sVal = Replace(sVal, "\""", """")
                sVal = Replace(sVal, "\\", "\")
            End If
            If Not dict.Exists(sKey) Then dict.Add sKey, sVal
        End If
    Next j

    Set SimpleJsonToDictionary = dict
End Function

' ===========================================================
'  REGEX HELPERS (uses VBScript.RegExp)
' ===========================================================

' Check if a regex pattern is valid
Public Function IsValidRegex(sPattern As String) As Boolean
    On Error Resume Next
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = sPattern
    Dim test As Boolean
    test = re.Test("")
    IsValidRegex = (Err.Number = 0)
    Set re = Nothing
    On Error GoTo 0
End Function

' Extract all matches of a regex from a string
Public Function RegexExtractAll(sInput As String, sPattern As String) As String()
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = sPattern
    re.Global = True
    re.IgnoreCase = False

    Dim result() As String
    ReDim result(0)
    Dim count As Integer: count = 0

    If re.Test(sInput) Then
        Dim matches As Object
        Set matches = re.Execute(sInput)
        Dim m As Object
        For Each m In matches
            ReDim Preserve result(count)
            result(count) = m.Value
            count = count + 1
        Next m
    End If

    Set re = Nothing
    If count = 0 Then
        ReDim result(0)
        result(0) = ""
    End If
    RegexExtractAll = result
End Function

' Test regex against sample data — returns matched text or "NO MATCH"
Public Function RegexTest(sPattern As String, sSample As String) As String
    If Not IsValidRegex(sPattern) Then
        RegexTest = "ERROR: Invalid pattern"
        Exit Function
    End If

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = sPattern
    re.Global = True

    If Not re.Test(sSample) Then
        RegexTest = "NO MATCH"
    Else
        Dim matches As Object
        Set matches = re.Execute(sSample)
        Dim result As String: result = ""
        Dim m As Object
        For Each m In matches
            If result <> "" Then result = result & ", "
            result = result & "[" & m.Value & "]"
        Next m
        RegexTest = result
    End If
    Set re = Nothing
End Function

' ===========================================================
'  STRING HELPERS
' ===========================================================

' Split a delimited string, handling quoted fields (CSV-style)
Public Function SplitDelimited(sLine As String, sDelim As String) As String()
    Dim result() As String
    ReDim result(0)
    Dim count As Integer: count = 0

    If sDelim = DELIM_COMMA Or sDelim = DELIM_TAB Or _
       sDelim = DELIM_PIPE Or sDelim = DELIM_SEMICOLON Then
        ' CSV-aware split
        Dim inQ As Boolean: inQ = False
        Dim token As String: token = ""
        Dim i As Integer
        For i = 1 To Len(sLine)
            Dim c As String: c = Mid(sLine, i, 1)
            If c = """" Then
                inQ = Not inQ
            ElseIf c = sDelim And Not inQ Then
                ReDim Preserve result(count)
                result(count) = Trim(token)
                count = count + 1
                token = ""
            Else
                token = token & c
            End If
        Next i
        ReDim Preserve result(count)
        result(count) = Trim(token)
    Else
        result = Split(sLine, sDelim)
    End If

    SplitDelimited = result
End Function

' Detect likely delimiter in a line of text
Public Function DetectDelimiter(sLine As String) As String
    Dim counts(4) As Integer
    Dim delims(4) As String
    delims(0) = DELIM_COMMA
    delims(1) = DELIM_TAB
    delims(2) = DELIM_PIPE
    delims(3) = DELIM_SEMICOLON
    delims(4) = DELIM_SPACE

    Dim i As Integer
    For i = 0 To 4
        counts(i) = Len(sLine) - Len(Replace(sLine, delims(i), ""))
    Next i

    Dim maxCount As Integer: maxCount = 0
    Dim bestDelim As String: bestDelim = DELIM_COMMA
    For i = 0 To 4
        If counts(i) > maxCount Then
            maxCount = counts(i)
            bestDelim = delims(i)
        End If
    Next i

    DetectDelimiter = bestDelim
End Function

' Clean a currency string to a Double ("$1,234.56" -> 1234.56)
Public Function ParseCurrencyString(s As String) As Double
    s = Trim(s)
    s = Replace(s, "$", "")
    s = Replace(s, ",", "")
    s = Replace(s, "(", "-")
    s = Replace(s, ")", "")
    On Error Resume Next
    ParseCurrencyString = CDbl(s)
    If Err.Number <> 0 Then ParseCurrencyString = 0
    On Error GoTo 0
End Function

' Generate a GUID-style unique ID
Public Function GenerateId() As String
    Dim d As Date: d = Now()
    Randomize
    GenerateId = Format(d, "yyyymmddhhmmss") & "_" & CStr(Int(Rnd() * 99999))
End Function

' Convert hex color string "#RRGGBB" to Excel Long
Public Function HexToLong(sHex As String) As Long
    sHex = Replace(sHex, "#", "")
    sHex = Replace(sHex, "0x", "")
    If Len(sHex) < 6 Then
        HexToLong = 0
        Exit Function
    End If
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Left(sHex, 2))
    g = CLng("&H" & Mid(sHex, 3, 2))
    b = CLng("&H" & Right(sHex, 2))
    ' Excel uses BGR order
    HexToLong = RGB(r, g, b)
End Function

' Convert Excel Long color to "#RRGGBB" hex string
Public Function LongToHex(lColor As Long) As String
    Dim r As Long, g As Long, b As Long
    b = lColor \ 65536
    g = (lColor Mod 65536) \ 256
    r = lColor Mod 256
    LongToHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

' Trim whitespace and strip non-printable chars from a cell value
Public Function CleanCellValue(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        CleanCellValue = ""
        Exit Function
    End If
    Dim s As String: s = CStr(v)
    s = Trim(s)
    ' Strip non-printable chars (< 32 except Tab and CR/LF)
    Dim i As Integer
    Dim result As String: result = ""
    For i = 1 To Len(s)
        Dim asc As Integer: asc = Asc(Mid(s, i, 1))
        If asc >= 32 Or asc = 9 Or asc = 10 Or asc = 13 Then
            result = result & Mid(s, i, 1)
        End If
    Next i
    CleanCellValue = Trim(result)
End Function

' ===========================================================
'  SHEET HELPERS
' ===========================================================

' Get a sheet by name; create it (hidden) if it doesn't exist
Public Function GetOrCreateSheet(sName As String, _
                                  Optional bHidden As Boolean = True) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sName
        If bHidden Then ws.Visible = xlSheetVeryHidden
    End If

    Set GetOrCreateSheet = ws
End Function

' Ensure a required hidden sheet exists with header row
Public Sub EnsureSystemSheet(sName As String, headers As Variant)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(sName, True)

    If ws.Cells(1, 1).Value = "" Then
        Dim i As Integer
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
            ws.Cells(1, i + 1).Font.Bold = True
        Next i
        ws.Rows(1).AutoFit
    End If
End Sub

' ===========================================================
'  FILE HELPERS
' ===========================================================

' Check if a file exists
Public Function FileExists(sPath As String) As Boolean
    FileExists = (Len(Dir(sPath)) > 0)
End Function

' Read first N bytes from a binary file (for magic byte detection)
Public Function ReadFileBytes(sPath As String, nBytes As Integer) As String
    Dim fileNum As Integer
    fileNum = FreeFile()
    Dim result As String
    result = String(nBytes, Chr(0))

    On Error Resume Next
    Open sPath For Binary Access Read As #fileNum
    If Err.Number = 0 Then
        Dim buf As String * 16
        Get #fileNum, , buf
        Close #fileNum
        result = Left(buf, nBytes)
    End If
    On Error GoTo 0

    ReadFileBytes = result
End Function

' Get temp folder path
Public Function GetTempPath() As String
    GetTempPath = Environ("TEMP")
    If Right(GetTempPath, 1) <> "\" Then GetTempPath = GetTempPath & "\"
End Function

' Get unique temp file path
Public Function GetTempFilePath(sExtension As String) As String
    GetTempFilePath = GetTempPath() & "RAE_" & Format(Now(), "yyyymmddhhmmss") & _
                      "_" & CStr(Int(Rnd() * 9999)) & "." & sExtension
End Function

' ===========================================================
'  SETTINGS HELPERS
' ===========================================================

Public Function GetSetting(sKey As String, sDefault As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SETTINGS)
    If ws Is Nothing Then
        GetSetting = sDefault
        Exit Function
    End If

    Dim i As Long
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(i, 1).Value) = sKey Then
            GetSetting = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    GetSetting = sDefault
    On Error GoTo 0
End Function

Public Sub SaveSetting(sKey As String, sValue As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_SETTINGS, True)
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Key"
        ws.Cells(1, 2).Value = "Value"
        ws.Cells(1, 1).Font.Bold = True
        ws.Cells(1, 2).Font.Bold = True
    End If

    ' Look for existing key
    Dim i As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = sKey Then
            ws.Cells(i, 2).Value = sValue
            Exit Sub
        End If
    Next i

    ' Not found — append
    ws.Cells(lastRow + 1, 1).Value = sKey
    ws.Cells(lastRow + 1, 2).Value = sValue
    On Error GoTo 0
End Sub

' ===========================================================
'  COLUMN LETTER / NUMBER CONVERSION
' ===========================================================

Public Function ColNumToLetter(n As Integer) As String
    If n <= 26 Then
        ColNumToLetter = Chr(64 + n)
    Else
        ColNumToLetter = Chr(64 + (n - 1) \ 26) & Chr(65 + (n - 1) Mod 26)
    End If
End Function

Public Function CellAddress(nRow As Integer, nCol As Integer) As String
    CellAddress = ColNumToLetter(nCol) & CStr(nRow)
End Function

' ===========================================================
'  MESSAGE BOX HELPERS (consistent styling)
' ===========================================================

Public Sub ShowInfo(sMsg As String, Optional sTitle As String = "")
    If sTitle = "" Then sTitle = APP_NAME
    MsgBox sMsg, vbInformation + vbOKOnly, sTitle
End Sub

Public Sub ShowError(sMsg As String, Optional sTitle As String = "")
    If sTitle = "" Then sTitle = APP_NAME & " — Error"
    MsgBox sMsg, vbCritical + vbOKOnly, sTitle
End Sub

Public Function ConfirmAction(sMsg As String, Optional sTitle As String = "") As Boolean
    If sTitle = "" Then sTitle = APP_NAME
    ConfirmAction = (MsgBox(sMsg, vbQuestion + vbYesNo, sTitle) = vbYes)
End Function
