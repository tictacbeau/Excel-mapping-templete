VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPayorManager
   Caption         =   "Payor Manager"
   ClientHeight    =   8000
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   10000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPayorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  frmPayorManager — List + CRUD for payor configurations
'  Displays all payors in a list; provides Add/Edit/Delete/Clone
' ============================================================

Private m_Payors As Collection
Private m_SelectedPayorId As String
Private m_SelectedPayorName As String

Private Sub UserForm_Initialize()
    lblTitle.Caption = "Payor Manager"
    lblTitle.Font.Bold = True
    lblTitle.Font.Size = 14
    RefreshPayorList
End Sub

Private Sub RefreshPayorList()
    lstPayors.Clear
    Set m_Payors = modPayorRegistry.GetAllPayors()

    If m_Payors.Count = 0 Then
        lblPayorCount.Caption = "No payors configured. Click 'Add New' to get started."
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnDuplicate.Enabled = False
        btnProcess.Enabled = False
        Exit Sub
    End If

    ' Populate list with payor names and summary info
    Dim oPayor As clsPayor
    For Each oPayor In m_Payors
        Dim sSingleStatus As String, sMultiStatus As String
        Dim oSingle As clsTemplate: Set oSingle = modPayorRegistry.LoadTemplate(oPayor.PayorId, TMPL_SINGLE)
        Dim oMulti As clsTemplate: Set oMulti = modPayorRegistry.LoadTemplate(oPayor.PayorId, TMPL_MULTI)
        sSingleStatus = IIf(Not oSingle Is Nothing And oSingle.CellCount > 0, "✓", "○")
        sMultiStatus = IIf(Not oMulti Is Nothing And oMulti.CellCount > 0, "✓", "○")

        Dim sLine As String
        sLine = oPayor.PayorName & _
            "  |  " & oPayor.ExpectedFileType & _
            "  |  Single:" & sSingleStatus & _
            "  Multi:" & sMultiStatus & _
            "  |  Modified: " & Format(oPayor.LastModified, "mm/dd/yy")
        lstPayors.AddItem sLine
    Next oPayor

    lblPayorCount.Caption = m_Payors.Count & " payor(s) configured."
    lstPayors.ListIndex = 0
    UpdateSelection
End Sub

Private Sub lstPayors_Click()
    UpdateSelection
End Sub

Private Sub UpdateSelection()
    Dim nIdx As Integer: nIdx = lstPayors.ListIndex
    If nIdx < 0 Or m_Payors Is Nothing Then
        m_SelectedPayorId = ""
        m_SelectedPayorName = ""
        btnEdit.Enabled = False
        btnDelete.Enabled = False
        btnDuplicate.Enabled = False
        btnProcess.Enabled = False
        lblSelectedInfo.Caption = ""
        Exit Sub
    End If

    ' Map list index to payor
    Dim i As Integer: i = 1
    Dim oPayor As clsPayor
    For Each oPayor In m_Payors
        If i - 1 = nIdx Then
            m_SelectedPayorId = oPayor.PayorId
            m_SelectedPayorName = oPayor.PayorName
            ShowPayorDetail oPayor
            Exit For
        End If
        i = i + 1
    Next oPayor

    btnEdit.Enabled = True
    btnDelete.Enabled = True
    btnDuplicate.Enabled = True
    btnProcess.Enabled = True
End Sub

Private Sub ShowPayorDetail(oPayor As clsPayor)
    Dim s As String
    s = "Payor: " & oPayor.PayorName & vbCrLf
    s = s & "ID: " & oPayor.PayorId & vbCrLf
    s = s & "File Type: " & oPayor.ExpectedFileType & vbCrLf
    s = s & "Invoice Layout: " & oPayor.InvoiceLayout & vbCrLf
    If oPayor.InvoiceRegex <> "" Then
        s = s & "Regex: " & oPayor.InvoiceRegex & vbCrLf
    End If
    s = s & "Duplicate Action: " & oPayor.DuplicateAction & vbCrLf
    s = s & "Last Modified: " & Format(oPayor.LastModified, "mm/dd/yyyy hh:mm") & vbCrLf
    If oPayor.Notes <> "" Then
        s = s & "Notes: " & oPayor.Notes & vbCrLf
    End If
    lblSelectedInfo.Caption = s
End Sub

' ===========================================================
'  ACTIONS
' ===========================================================

Private Sub btnAddNew_Click()
    Me.Hide
    modMain.Open_NewPayorWizard
    RefreshPayorList
    Me.Show
End Sub

Private Sub btnEdit_Click()
    If m_SelectedPayorId = "" Then Exit Sub
    Me.Hide
    modMain.Edit_Payor m_SelectedPayorId
    RefreshPayorList
    Me.Show
End Sub

Private Sub btnDelete_Click()
    If m_SelectedPayorId = "" Then Exit Sub
    If Not modUtils.ConfirmAction( _
        "Delete payor '" & m_SelectedPayorName & "'?" & vbCrLf & _
        "All associated templates will also be deleted. This cannot be undone.") Then Exit Sub
    modPayorRegistry.DeletePayor m_SelectedPayorId
    RefreshPayorList
End Sub

Private Sub btnDuplicate_Click()
    If m_SelectedPayorId = "" Then Exit Sub

    Dim sNewName As String
    sNewName = InputBox( _
        "Enter a name for the new payor (copy of '" & m_SelectedPayorName & "'):", _
        "Duplicate Payor", m_SelectedPayorName & " (Copy)")

    If sNewName = "" Then Exit Sub

    If modPayorRegistry.PayorNameExists(sNewName) Then
        MsgBox "A payor named '" & sNewName & "' already exists.", vbExclamation
        Exit Sub
    End If

    Dim oNew As clsPayor
    Set oNew = modPayorRegistry.DuplicatePayor(m_SelectedPayorId, sNewName)
    If Not oNew Is Nothing Then
        MsgBox "Payor '" & sNewName & "' created successfully.", vbInformation
        RefreshPayorList
    Else
        MsgBox "Duplication failed. Check audit log.", vbCritical
    End If
End Sub

Private Sub btnProcess_Click()
    If m_SelectedPayorId = "" Then Exit Sub
    Me.Hide
    modMain.Process_Remittance
End Sub

Private Sub btnExport_Click()
    If m_SelectedPayorId = "" Then Exit Sub
    modMain.ExportPayorConfig m_SelectedPayorId
End Sub

Private Sub btnImport_Click()
    modMain.ImportPayorConfig
    RefreshPayorList
End Sub

Private Sub btnViewTemplates_Click()
    If m_SelectedPayorId = "" Then Exit Sub
    Dim oPayor As clsPayor
    Set oPayor = modPayorRegistry.GetPayor(m_SelectedPayorId)
    If oPayor Is Nothing Then Exit Sub
    Me.Hide
    modMain.Edit_Payor m_SelectedPayorId
    RefreshPayorList
    Me.Show
End Sub

Private Sub btnRefreshList_Click()
    RefreshPayorList
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub lstPayors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnEdit_Click
End Sub

' ---- Controls summary: ----
' lblTitle          — "Payor Manager"
' lstPayors         — List box (payor list with summary)
' lblPayorCount     — "N payor(s) configured"
' lblSelectedInfo   — Detail panel for selected payor
' btnAddNew         — Add new payor (opens wizard)
' btnEdit           — Edit selected payor (opens wizard)
' btnDelete         — Delete selected payor
' btnDuplicate      — Clone selected payor
' btnProcess        — Process remittance for selected payor
' btnExport         — Export selected payor config as JSON
' btnImport         — Import payor config from JSON
' btnViewTemplates  — View/edit templates for selected payor
' btnRefreshList    — Refresh the payor list
' btnClose          — Close the manager
