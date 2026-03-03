VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProcessRemittance
   Caption         =   "Process Remittance File"
   ClientHeight    =   7500
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   9600
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProcessRemittance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  frmProcessRemittance — Remittance Processing Dialog
'  Step 1: Select Payor
'  Step 2: Select File
'  Step 3: Preview / Confirm
'  Step 4: Results
' ============================================================

Private m_SelectedPayorId As String
Private m_SelectedFile As String
Private m_SelectedPayorName As String
Private m_CurrentStep As Integer

Private Sub UserForm_Initialize()
    m_CurrentStep = 1
    LoadPayorDropdown
    UpdateStepDisplay
    lblTitle.Caption = "Process Remittance"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 14
End Sub

Private Sub LoadPayorDropdown()
    cboPayor.Clear
    Dim names() As String
    names = modPayorRegistry.GetPayorNameList()
    Dim i As Integer
    For i = 0 To UBound(names)
        If names(i) <> "" Then cboPayor.AddItem names(i)
    Next i
    If cboPayor.ListCount > 0 Then
        ' Try to restore last used payor
        Dim sLast As String
        sLast = modUtils.GetSetting(SETTING_LAST_PAYOR, "")
        If sLast <> "" Then
            cboPayor.Text = sLast
        Else
            cboPayor.ListIndex = 0
        End If
        UpdatePayorInfo
    End If
End Sub

Private Sub cboPayor_Change()
    UpdatePayorInfo
End Sub

Private Sub UpdatePayorInfo()
    m_SelectedPayorName = cboPayor.Text
    Dim oPayor As clsPayor
    Set oPayor = modPayorRegistry.GetPayorByName(m_SelectedPayorName)
    If oPayor Is Nothing Then
        lblPayorInfo.Caption = "Payor not found."
        m_SelectedPayorId = ""
    Else
        m_SelectedPayorId = oPayor.PayorId
        lblPayorInfo.Caption = _
            "Expected file type: " & oPayor.ExpectedFileType & vbCrLf & _
            "Invoice layout: " & oPayor.InvoiceLayout & vbCrLf & _
            "Regex: " & IIf(oPayor.InvoiceRegex <> "", oPayor.InvoiceRegex, "(none)")
    End If
End Sub

Private Sub btnBrowse_Click()
    Dim sFilter As String
    sFilter = "Remittance Files (*.xlsx;*.xls;*.csv;*.txt),*.xlsx;*.xls;*.csv;*.txt"

    ' Suggest expected type based on payor
    Dim oPayor As clsPayor
    If m_SelectedPayorId <> "" Then
        Set oPayor = modPayorRegistry.GetPayor(m_SelectedPayorId)
    End If

    Dim sFile As String
    sFile = Application.GetOpenFilename(sFilter, 1, _
        "Select Remittance File" & IIf(m_SelectedPayorName <> "", " for " & m_SelectedPayorName, ""), _
        , False)
    If sFile <> "False" And sFile <> "" Then
        m_SelectedFile = sFile
        txtFilePath.Value = sFile
        lblFileStatus.Caption = "File selected: " & Right(sFile, 50)
        lblFileStatus.ForeColor = RGB(0, 128, 0)
        UpdateStepDisplay
    End If
End Sub

Private Sub btnProcess_Click()
    If m_SelectedPayorId = "" Then
        MsgBox "Please select a payor.", vbExclamation, APP_NAME
        Exit Sub
    End If
    If m_SelectedFile = "" Then
        MsgBox "Please select a remittance file.", vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not modUtils.FileExists(m_SelectedFile) Then
        MsgBox "Selected file no longer exists: " & m_SelectedFile, vbCritical, APP_NAME
        Exit Sub
    End If

    ' Save last payor
    modUtils.SaveSetting SETTING_LAST_PAYOR, m_SelectedPayorName

    ' Disable UI and show progress
    btnProcess.Enabled = False
    lblStatus.Caption = "Processing... please wait."
    lblStatus.ForeColor = RGB(0, 0, 200)
    Me.Repaint
    Application.StatusBar = "Processing remittance for " & m_SelectedPayorName & "..."

    ' Run the engine
    Dim wsResult As Worksheet
    Set wsResult = modTemplateEngine.ProcessRemittance(m_SelectedPayorId, m_SelectedFile)

    Application.StatusBar = False
    btnProcess.Enabled = True

    If Not wsResult Is Nothing Then
        lblStatus.Caption = "Processing complete! Output: " & wsResult.Name
        lblStatus.ForeColor = RGB(0, 128, 0)
        MsgBox "Remittance processed successfully!" & vbCrLf & _
               "Output worksheet: " & wsResult.Name, vbInformation, APP_NAME
        Me.Hide
    Else
        lblStatus.Caption = "Processing failed. See audit log for details."
        lblStatus.ForeColor = RGB(200, 0, 0)
    End If
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnNewPayor_Click()
    Me.Hide
    modMain.Open_NewPayorWizard
End Sub

Private Sub UpdateStepDisplay()
    Dim bReady As Boolean
    bReady = (m_SelectedPayorId <> "" And m_SelectedFile <> "")
    btnProcess.Enabled = bReady
    If bReady Then
        lblReadyStatus.Caption = "Ready to process!"
        lblReadyStatus.ForeColor = RGB(0, 128, 0)
    Else
        lblReadyStatus.Caption = "Select a payor and file to continue."
        lblReadyStatus.ForeColor = RGB(100, 100, 100)
    End If
End Sub

' ---- Controls defined in frm layout: ----
' lblTitle          — Title label
' cboPayor          — Payor dropdown
' lblPayorInfo      — Payor details text
' txtFilePath       — File path text box
' btnBrowse         — Browse button
' lblFileStatus     — File status message
' btnProcess        — Process button (default)
' btnCancel         — Cancel button
' btnNewPayor       — "Set up new payor" shortcut
' lblStatus         — Processing status message
' lblReadyStatus    — Ready state indicator
