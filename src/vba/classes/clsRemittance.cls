VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRemittance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  clsRemittance — Parsed remittance data from one payor file
'  After file conversion and parsing, this object holds all
'  extracted invoice data ready for template population.
' ============================================================

' ---- Source Info -------------------------------------------
Public SourceFilePath As String       ' Original file path (pre-conversion)
Public NormalizedFilePath As String   ' Post-conversion .xlsx path
Public SourceFileType As String       ' Detected type: xlsx/xls_genuine/xls_html/csv
Public PayorId As String              ' Associated payor
Public PayorName As String
Public ImportedAt As Date

' ---- Payment Header Data -----------------------------------
Public PaymentDate As String          ' As text, formatted by payor
Public PaymentAmount As Double        ' Total remittance amount
Public CheckNumber As String          ' Check or EFT reference number
Public RemittanceId As String         ' Any remittance/batch ID
Public RemittanceNotes As String      ' Free-text notes from payor

' ---- Invoice Collection ------------------------------------
' Parallel arrays: one entry per extracted invoice
Private m_InvoiceNumbers() As String
Private m_InvoiceAmounts() As Double
Private m_InvoiceRawData() As String  ' Full raw text for that invoice's row
Private m_InvoiceDates() As String
Private m_InvoiceDuplicateFlag() As Boolean
Private m_Count As Integer

' ---- Raw Worksheet Reference (temp, after conversion) ------
Public RawWorkbook As Workbook        ' The normalized intermediate workbook

' ---- Audit / Conversion Log --------------------------------
Private m_AuditLines() As String
Private m_AuditCount As Integer

' ---- Parse Warnings ----------------------------------------
Private m_Warnings() As String
Private m_WarnCount As Integer

' ===========================================================
'  Initialize
' ===========================================================
Private Sub Class_Initialize()
    m_Count = 0
    ReDim m_InvoiceNumbers(0)
    ReDim m_InvoiceAmounts(0)
    ReDim m_InvoiceRawData(0)
    ReDim m_InvoiceDates(0)
    ReDim m_InvoiceDuplicateFlag(0)
    m_AuditCount = 0
    ReDim m_AuditLines(0)
    m_WarnCount = 0
    ReDim m_Warnings(0)
    ImportedAt = Now()
    PaymentAmount = 0
    CheckNumber = ""
    RemittanceId = ""
    SourceFileType = FILETYPE_XLSX
End Sub

' ===========================================================
'  Add one invoice to the collection
' ===========================================================
Public Sub AddInvoice(sInvoiceNo As String, dAmount As Double, _
                       Optional sDate As String = "", _
                       Optional sRawRow As String = "")
    ' Check for duplicate
    Dim bDuplicate As Boolean: bDuplicate = False
    Dim i As Integer
    For i = 0 To m_Count - 1
        If m_InvoiceNumbers(i) = Trim(sInvoiceNo) Then
            bDuplicate = True
            Exit For
        End If
    Next i

    ReDim Preserve m_InvoiceNumbers(m_Count)
    ReDim Preserve m_InvoiceAmounts(m_Count)
    ReDim Preserve m_InvoiceRawData(m_Count)
    ReDim Preserve m_InvoiceDates(m_Count)
    ReDim Preserve m_InvoiceDuplicateFlag(m_Count)

    m_InvoiceNumbers(m_Count) = Trim(sInvoiceNo)
    m_InvoiceAmounts(m_Count) = dAmount
    m_InvoiceRawData(m_Count) = sRawRow
    m_InvoiceDates(m_Count) = sDate
    m_InvoiceDuplicateFlag(m_Count) = bDuplicate

    If bDuplicate Then
        AddWarning "Duplicate invoice detected: " & sInvoiceNo & " (occurrence " & (m_Count + 1) & ")"
    End If

    m_Count = m_Count + 1
End Sub

' ===========================================================
'  Read-only access to invoice data
' ===========================================================
Public Property Get InvoiceCount() As Integer
    InvoiceCount = m_Count
End Property

Public Property Get InvoiceNumber(i As Integer) As String
    If i >= 0 And i < m_Count Then
        InvoiceNumber = m_InvoiceNumbers(i)
    Else
        InvoiceNumber = ""
    End If
End Property

Public Property Get InvoiceAmount(i As Integer) As Double
    If i >= 0 And i < m_Count Then
        InvoiceAmount = m_InvoiceAmounts(i)
    Else
        InvoiceAmount = 0
    End If
End Property

Public Property Get InvoiceDate(i As Integer) As String
    If i >= 0 And i < m_Count Then
        InvoiceDate = m_InvoiceDates(i)
    Else
        InvoiceDate = ""
    End If
End Property

Public Property Get InvoiceRawRow(i As Integer) As String
    If i >= 0 And i < m_Count Then
        InvoiceRawRow = m_InvoiceRawData(i)
    Else
        InvoiceRawRow = ""
    End If
End Property

Public Property Get IsDuplicate(i As Integer) As Boolean
    If i >= 0 And i < m_Count Then
        IsDuplicate = m_InvoiceDuplicateFlag(i)
    Else
        IsDuplicate = False
    End If
End Property

' ===========================================================
'  Calculated total of all invoice amounts
' ===========================================================
Public Property Get TotalExtracted() As Double
    Dim d As Double: d = 0
    Dim i As Integer
    For i = 0 To m_Count - 1
        d = d + m_InvoiceAmounts(i)
    Next i
    TotalExtracted = d
End Property

' ===========================================================
'  Variance between remittance total and extracted sum
' ===========================================================
Public Property Get Variance() As Double
    Variance = PaymentAmount - TotalExtracted
End Property

' ===========================================================
'  Is single-invoice?
' ===========================================================
Public Property Get IsSingleInvoice() As Boolean
    IsSingleInvoice = (m_Count = 1)
End Property

' ===========================================================
'  Audit log
' ===========================================================
Public Sub AddAuditLine(sLine As String)
    ReDim Preserve m_AuditLines(m_AuditCount)
    m_AuditLines(m_AuditCount) = Format(Now(), "hh:mm:ss") & " - " & sLine
    m_AuditCount = m_AuditCount + 1
End Sub

Public Property Get AuditLog() As String
    AuditLog = Join(m_AuditLines, vbCrLf)
End Property

Public Property Get AuditLineCount() As Integer
    AuditLineCount = m_AuditCount
End Property

' ===========================================================
'  Warnings
' ===========================================================
Public Sub AddWarning(sMsg As String)
    ReDim Preserve m_Warnings(m_WarnCount)
    m_Warnings(m_WarnCount) = sMsg
    m_WarnCount = m_WarnCount + 1
End Sub

Public Property Get WarningCount() As Integer
    WarningCount = m_WarnCount
End Property

Public Property Get Warning(i As Integer) As String
    If i >= 0 And i < m_WarnCount Then
        Warning = m_Warnings(i)
    Else
        Warning = ""
    End If
End Property

Public Property Get AllWarnings() As String
    AllWarnings = Join(m_Warnings, vbCrLf)
End Property

' ===========================================================
'  Has any duplicate invoices?
' ===========================================================
Public Property Get HasDuplicates() As Boolean
    Dim i As Integer
    For i = 0 To m_Count - 1
        If m_InvoiceDuplicateFlag(i) Then
            HasDuplicates = True
            Exit Property
        End If
    Next i
    HasDuplicates = False
End Property

' ===========================================================
'  Serialize summary to string for audit trail
' ===========================================================
Public Function ToSummaryString() As String
    Dim s As String
    s = "Payor: " & PayorName & vbCrLf
    s = s & "File: " & SourceFilePath & vbCrLf
    s = s & "File Type: " & SourceFileType & vbCrLf
    s = s & "Invoice Count: " & m_Count & vbCrLf
    s = s & "Template Type: " & IIf(IsSingleInvoice, "SINGLE", "MULTI") & vbCrLf
    s = s & "Remittance Total: " & Format(PaymentAmount, "$#,##0.00") & vbCrLf
    s = s & "Extracted Total: " & Format(TotalExtracted, "$#,##0.00") & vbCrLf
    s = s & "Variance: " & Format(Variance, "$#,##0.00") & vbCrLf
    If HasDuplicates Then s = s & "WARNING: Duplicate invoice numbers detected." & vbCrLf
    ToSummaryString = s
End Function

' ===========================================================
'  Cleanup
' ===========================================================
Private Sub Class_Terminate()
    On Error Resume Next
    If Not RawWorkbook Is Nothing Then
        RawWorkbook.Close SaveChanges:=False
    End If
    Set RawWorkbook = Nothing
    On Error GoTo 0
End Sub
