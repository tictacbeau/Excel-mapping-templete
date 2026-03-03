Attribute VB_Name = "modDataLookup"
Option Explicit

' ============================================================
'  modDataLookup — Power Query + Data Model Integration
'
'  Manages the internal data lookup system:
'    - Connects to and refreshes the Power Query data source
'    - Queries the Data Model (Power Pivot) for invoice data
'    - Returns lookup results as Scripting.Dictionary objects
'    - Validates schema drift on data refresh
'    - Provides rollback to previous data snapshot
' ============================================================

' Module-level cache (Dictionary of Dictionaries: invoiceNo -> field values)
Private m_LookupCache As Object        ' Scripting.Dictionary
Private m_CacheLoaded As Boolean
Private m_CachePayorId As String       ' Cache is payor-independent (global internal data)

' Expected schema (column names) — set on first load, checked on refresh
Private m_SchemaColumns() As String
Private m_SchemaCount As Integer

' ===========================================================
'  LOOKUP ONE INVOICE in Data Model
'  Returns: Scripting.Dictionary (field -> value) or Nothing
' ===========================================================
Public Function LookupInvoice(sInvoiceNo As String, _
                               oConfig As Object) As Object
    If sInvoiceNo = "" Then
        Set LookupInvoice = Nothing
        Exit Function
    End If

    ' Ensure cache is loaded
    If Not m_CacheLoaded Then
        LoadLookupCache oConfig
    End If

    ' Try exact match first
    Dim sKey As String: sKey = Trim(sInvoiceNo)
    If m_LookupCache.Exists(sKey) Then
        Set LookupInvoice = m_LookupCache(sKey)
        Exit Function
    End If

    ' Check match rules
    Dim sMatchType As String: sMatchType = "exact"
    If Not oConfig Is Nothing Then
        If oConfig.Exists("matchType") Then sMatchType = LCase(oConfig("matchType"))
    End If

    If sMatchType = "contains" Or sMatchType = "partial" Then
        ' Scan cache for partial match
        Dim key As Variant
        For Each key In m_LookupCache.Keys
            If InStr(LCase(CStr(key)), LCase(sKey)) > 0 Or _
               InStr(LCase(sKey), LCase(CStr(key))) > 0 Then
                Set LookupInvoice = m_LookupCache(key)
                Exit Function
            End If
        Next key
    End If

    Set LookupInvoice = Nothing
End Function

' ===========================================================
'  LOAD LOOKUP CACHE from Data Model via Power Query
'  Reads the InternalData table from the Data Model worksheet
'  and populates the in-memory Dictionary cache.
' ===========================================================
Public Sub LoadLookupCache(Optional oConfig As Object = Nothing)
    Set m_LookupCache = CreateObject("Scripting.Dictionary")
    m_LookupCache.CompareMode = vbTextCompare
    m_CacheLoaded = False

    ' Find the data connection/table
    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()

    If wsData Is Nothing Then
        modUtils.WriteAuditLog "WARNING", "DataLookup", _
            "No internal data table found. Run Data Model Refresh to import data.", , "", "WARN"
        m_CacheLoaded = True  ' Loaded (empty)
        Exit Sub
    End If

    ' Determine key column
    Dim sKeyField As String: sKeyField = "InvoiceNumber"
    If Not oConfig Is Nothing Then
        If oConfig.Exists("primaryKey") Then sKeyField = oConfig("primaryKey")
    End If

    ' Read headers from row 1
    Dim nLastCol As Integer
    nLastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    If nLastCol < 1 Or wsData.Cells(1, 1).Value = "" Then
        modUtils.WriteAuditLog "WARNING", "DataLookup", _
            "Internal data table appears empty.", , "", "WARN"
        m_CacheLoaded = True
        Exit Sub
    End If

    Dim headers() As String
    ReDim headers(nLastCol - 1)
    Dim nKeyCol As Integer: nKeyCol = -1
    Dim c As Integer
    For c = 1 To nLastCol
        headers(c - 1) = Trim(wsData.Cells(1, c).Value)
        If LCase(headers(c - 1)) = LCase(sKeyField) Then nKeyCol = c
    Next c

    If nKeyCol = -1 Then
        ' Try to find invoice-related column by common names
        Dim commonNames(4) As String
        commonNames(0) = "InvoiceNumber": commonNames(1) = "InvoiceNo"
        commonNames(2) = "Invoice": commonNames(3) = "InvNumber": commonNames(4) = "InvNo"
        For c = 1 To nLastCol
            Dim cn As Integer
            For cn = 0 To 4
                If LCase(headers(c - 1)) = LCase(commonNames(cn)) Then
                    nKeyCol = c
                    sKeyField = headers(c - 1)
                    Exit For
                End If
            Next cn
            If nKeyCol > -1 Then Exit For
        Next c
    End If

    If nKeyCol = -1 Then
        modUtils.WriteAuditLog "WARNING", "DataLookup", _
            "Cannot find key column '" & sKeyField & "' in internal data table.", , "", "WARN"
        m_CacheLoaded = True
        Exit Sub
    End If

    ' Cache schema
    m_SchemaCount = nLastCol
    ReDim m_SchemaColumns(nLastCol - 1)
    For c = 0 To nLastCol - 1: m_SchemaColumns(c) = headers(c): Next c

    ' Read all data rows into cache
    Dim nLastRow As Long
    nLastRow = wsData.Cells(wsData.Rows.Count, nKeyCol).End(xlUp).Row

    Dim i As Long
    Dim loaded As Long: loaded = 0
    For i = 2 To nLastRow
        Dim sKey As String: sKey = Trim(modUtils.CleanCellValue(wsData.Cells(i, nKeyCol).Value))
        If sKey = "" Then GoTo NextCacheRow

        Dim oRow As Object
        Set oRow = CreateObject("Scripting.Dictionary")
        For c = 1 To nLastCol
            oRow(headers(c - 1)) = modUtils.CleanCellValue(wsData.Cells(i, c).Value)
        Next c

        If Not m_LookupCache.Exists(sKey) Then
            m_LookupCache.Add sKey, oRow
            loaded = loaded + 1
        End If
NextCacheRow:
    Next i

    m_CacheLoaded = True
    modUtils.WriteAuditLog "LOAD_CACHE", "DataLookup", _
        "Loaded " & loaded & " records from internal data table.", , "", "OK"
End Sub

' ===========================================================
'  REFRESH DATA MODEL from a new source file
'  This is the full import/refresh wizard backend.
' ===========================================================
Public Function RefreshDataModel(sSourceFile As String, _
                                   Optional sKeyColumn As String = "InvoiceNumber") As Boolean
    RefreshDataModel = False

    If Not modUtils.FileExists(sSourceFile) Then
        modUtils.ShowError "Source file not found: " & sSourceFile
        Exit Function
    End If

    ' Step 1: Backup current data
    Dim nBefore As Long: nBefore = 0
    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()
    If Not wsData Is Nothing Then
        nBefore = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row - 1
        BackupDataModelSheet wsData
    End If

    ' Step 2: Determine source file type and open it
    Dim wbSource As Workbook
    On Error GoTo RefreshErr
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim sExt As String
    sExt = LCase(Right(sSourceFile, Len(sSourceFile) - InStrRev(sSourceFile, ".")))

    Dim oTempRemittance As New clsRemittance
    Dim sNormalized As String

    If sExt = "csv" Or sExt = "txt" Then
        sNormalized = modFileConverter.ConvertToXlsx(sSourceFile, oTempRemittance)
        If sNormalized = "" Then GoTo RefreshErr
        Set wbSource = Application.Workbooks.Open(sNormalized, ReadOnly:=True)
    Else
        sNormalized = ""
        Set wbSource = Application.Workbooks.Open(sSourceFile, ReadOnly:=True)
    End If

    ' Step 3: Get or create the data model sheet
    If wsData Is Nothing Then
        Set wsData = modUtils.GetOrCreateSheet("_InternalData_", True)
    Else
        ' Clear existing data (but keep header)
        If wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row > 1 Then
            wsData.Rows("2:" & wsData.Rows.Count).Delete
        End If
    End If

    ' Step 4: Copy data from source to data model sheet
    Dim wsSource As Worksheet
    Set wsSource = wbSource.Sheets(1)
    Dim nSrcLastRow As Long: nSrcLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    Dim nSrcLastCol As Integer: nSrcLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    If nSrcLastRow < 1 Or nSrcLastCol < 1 Then
        wbSource.Close False
        modUtils.ShowError "Source file appears to be empty."
        GoTo RefreshErr
    End If

    ' Detect schema drift
    Dim nAfterCols As Integer: nAfterCols = nSrcLastCol
    Dim schemaWarnings As String: schemaWarnings = ""
    If m_SchemaCount > 0 Then
        Dim newHeaders() As String
        ReDim newHeaders(nSrcLastCol - 1)
        Dim c As Integer
        For c = 1 To nSrcLastCol
            newHeaders(c - 1) = Trim(wsSource.Cells(1, c).Value)
        Next c
        schemaWarnings = DetectSchemaDrift(m_SchemaColumns, newHeaders)
    End If

    ' Copy all data
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(nSrcLastRow, nSrcLastCol)).Copy
    wsData.Cells(1, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    wbSource.Close SaveChanges:=False
    If sNormalized <> "" Then
        On Error Resume Next: Kill sNormalized: On Error GoTo RefreshErr
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Step 5: Clear cache so it's reloaded on next lookup
    InvalidateCache

    ' Step 6: Validate and report
    Dim nAfter As Long: nAfter = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row - 1
    Dim sReport As String
    sReport = "Data refresh complete." & vbCrLf & _
              "Records before: " & nBefore & vbCrLf & _
              "Records after: " & nAfter & vbCrLf & _
              "Change: " & (nAfter - nBefore)

    If schemaWarnings <> "" Then
        sReport = sReport & vbCrLf & vbCrLf & "SCHEMA WARNINGS:" & vbCrLf & schemaWarnings
    End If

    modUtils.WriteAuditLog "REFRESH_DATA", "DataLookup", _
        "Refreshed " & nAfter & " records from: " & sSourceFile, , sSourceFile, "OK"
    modUtils.SaveSetting SETTING_DATA_SOURCE_PATH, sSourceFile

    RefreshDataModel = True
    Exit Function

RefreshErr:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close False
    On Error GoTo 0
    modUtils.WriteAuditLog "ERROR", "DataLookup.RefreshDataModel", Err.Description, , sSourceFile, "FAIL"
    Err.Clear
End Function

' ===========================================================
'  ROLLBACK to previous backup
' ===========================================================
Public Function RollbackDataModel() As Boolean
    RollbackDataModel = False
    Dim wsBackup As Worksheet
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Sheets("_InternalData_BACKUP")
    On Error GoTo 0

    If wsBackup Is Nothing Then
        modUtils.ShowError "No backup found to roll back to."
        Exit Function
    End If

    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()
    If wsData Is Nothing Then
        modUtils.ShowError "Internal data sheet not found."
        Exit Function
    End If

    ' Clear current and copy backup data
    wsData.Cells.Clear
    wsBackup.UsedRange.Copy
    wsData.Cells(1, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    InvalidateCache
    modUtils.WriteAuditLog "ROLLBACK", "DataLookup", "Data model rolled back to backup.", , "", "OK"
    RollbackDataModel = True
End Function

' ===========================================================
'  VALIDATE SCHEMA of current data source
'  Returns a human-readable report string.
' ===========================================================
Public Function ValidateDataModelSchema() As String
    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()
    If wsData Is Nothing Then
        ValidateDataModelSchema = "No internal data table found."
        Exit Function
    End If

    Dim nLastCol As Integer
    nLastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    Dim nLastRow As Long
    nLastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim report As String
    report = "Internal Data Model Schema:" & vbCrLf
    report = report & "  Columns: " & nLastCol & vbCrLf
    report = report & "  Rows: " & (nLastRow - 1) & vbCrLf
    report = report & "  Columns:" & vbCrLf

    Dim c As Integer
    For c = 1 To nLastCol
        Dim sHdr As String: sHdr = Trim(wsData.Cells(1, c).Value)
        report = report & "    " & c & ". " & sHdr & vbCrLf
    Next c

    ValidateDataModelSchema = report
End Function

' ===========================================================
'  GET RECORD COUNT in data model
' ===========================================================
Public Function GetDataModelRecordCount() As Long
    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()
    If wsData Is Nothing Then
        GetDataModelRecordCount = 0
        Exit Function
    End If
    Dim n As Long: n = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    GetDataModelRecordCount = IIf(n > 1, n - 1, 0)
End Function

' ===========================================================
'  GET DATA MODEL COLUMN NAMES for mapping UI
' ===========================================================
Public Function GetDataModelColumns() As String()
    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()
    Dim result() As String
    ReDim result(0)
    If wsData Is Nothing Then
        GetDataModelColumns = result
        Exit Function
    End If

    Dim nLastCol As Integer
    nLastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    ReDim result(nLastCol - 1)
    Dim c As Integer
    For c = 1 To nLastCol
        result(c - 1) = Trim(wsData.Cells(1, c).Value)
    Next c
    GetDataModelColumns = result
End Function

' ===========================================================
'  INVALIDATE the in-memory cache (forces reload on next use)
' ===========================================================
Public Sub InvalidateCache()
    Set m_LookupCache = Nothing
    m_CacheLoaded = False
End Sub

' ===========================================================
'  IS DATA MODEL LOADED?
' ===========================================================
Public Property Get IsDataModelLoaded() As Boolean
    Dim wsData As Worksheet
    Set wsData = FindDataModelSheet()
    If wsData Is Nothing Then
        IsDataModelLoaded = False
        Exit Property
    End If
    IsDataModelLoaded = (wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row > 1)
End Property

' ===========================================================
'  PRIVATE HELPERS
' ===========================================================

' Find the internal data storage sheet
Private Function FindDataModelSheet() As Worksheet
    On Error Resume Next
    Set FindDataModelSheet = ThisWorkbook.Sheets("_InternalData_")
    On Error GoTo 0
End Function

' Backup the current data model sheet
Private Sub BackupDataModelSheet(wsData As Worksheet)
    On Error Resume Next
    ' Delete existing backup
    Dim wsOldBackup As Worksheet
    Set wsOldBackup = ThisWorkbook.Sheets("_InternalData_BACKUP")
    If Not wsOldBackup Is Nothing Then
        Application.DisplayAlerts = False
        wsOldBackup.Delete
        Application.DisplayAlerts = True
    End If

    ' Copy current data to backup sheet
    wsData.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "_InternalData_BACKUP"
    ThisWorkbook.Sheets("_InternalData_BACKUP").Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' Compare two column name arrays and return drift description
Private Function DetectSchemaDrift(oldCols() As String, newCols() As String) As String
    Dim warnings As String: warnings = ""
    Dim i As Integer, j As Integer
    Dim found As Boolean

    ' Check for removed columns
    For i = 0 To UBound(oldCols)
        found = False
        For j = 0 To UBound(newCols)
            If LCase(oldCols(i)) = LCase(newCols(j)) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            warnings = warnings & "  REMOVED: " & oldCols(i) & vbCrLf
        End If
    Next i

    ' Check for added columns
    For j = 0 To UBound(newCols)
        found = False
        For i = 0 To UBound(oldCols)
            If LCase(newCols(j)) = LCase(oldCols(i)) Then
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            warnings = warnings & "  ADDED: " & newCols(j) & vbCrLf
        End If
    Next j

    DetectSchemaDrift = warnings
End Function

' ===========================================================
'  POWER QUERY REFRESH (for external .xlsx data source)
'  If a Power Query connection named PQ_CONNECTION_NAME exists,
'  refresh it. Otherwise use the direct import approach.
' ===========================================================
Public Sub RefreshPowerQueryConnection()
    On Error Resume Next
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        If conn.Name = PQ_CONNECTION_NAME Then
            conn.Refresh
            modUtils.WriteAuditLog "PQ_REFRESH", "DataLookup", _
                "Power Query connection refreshed: " & PQ_CONNECTION_NAME, , "", "OK"
            Exit For
        End If
    Next conn
    On Error GoTo 0
End Sub

' ===========================================================
'  CREATE POWER QUERY CONNECTION to a source file
'  (Excel 2016+ only; uses WorkbookQuery object)
' ===========================================================
Public Function CreatePowerQueryConnection(sSourceFile As String, _
                                             sTableName As String) As Boolean
    CreatePowerQueryConnection = False
    On Error GoTo PQErr

    ' Build M formula for a simple file import
    Dim sExt As String
    sExt = LCase(Right(sSourceFile, Len(sSourceFile) - InStrRev(sSourceFile, ".")))
    Dim sMFormula As String

    Select Case sExt
        Case "xlsx", "xlsm"
            sMFormula = "let" & vbCrLf & _
                "    Source = Excel.Workbook(File.Contents(""" & sSourceFile & """), null, true)," & vbCrLf & _
                "    Sheet1 = Source{[Item=""" & sTableName & """,Kind=""Sheet""]}[Data]," & vbCrLf & _
                "    PromotedHeaders = Table.PromoteHeaders(Sheet1, [PromoteAllScalars=true])," & vbCrLf & _
                "    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {})" & vbCrLf & _
                "in" & vbCrLf & _
                "    ChangedTypes"

        Case "csv"
            sMFormula = "let" & vbCrLf & _
                "    Source = Csv.Document(File.Contents(""" & sSourceFile & """), [Delimiter="","", Encoding=65001])," & vbCrLf & _
                "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & vbCrLf & _
                "in" & vbCrLf & _
                "    PromotedHeaders"

        Case Else
            modUtils.ShowError "Unsupported file type for Power Query: " & sExt
            Exit Function
    End Select

    ' Remove existing query with same name
    On Error Resume Next
    Dim q As WorkbookQuery
    For Each q In ThisWorkbook.Queries
        If q.Name = PQ_CONNECTION_NAME Then
            q.Delete
            Exit For
        End If
    Next q
    On Error GoTo PQErr

    ' Add the new query
    ThisWorkbook.Queries.Add PQ_CONNECTION_NAME, sMFormula

    ' Refresh connections
    RefreshPowerQueryConnection

    CreatePowerQueryConnection = True
    Exit Function

PQErr:
    modUtils.WriteAuditLog "ERROR", "DataLookup.CreatePowerQueryConnection", _
        Err.Description, , sSourceFile, "FAIL"
    CreatePowerQueryConnection = False
    Err.Clear
End Function
