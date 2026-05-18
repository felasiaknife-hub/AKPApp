VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListInv 
   Caption         =   "List Invoice"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9345
      ScaleWidth      =   2625
      TabIndex        =   8
      Top             =   0
      Width           =   2655
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2325
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmListInv.frx":0000
         Left            =   180
         List            =   "frmListInv.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1500
         Width           =   2325
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   5
         Top             =   4140
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   6
         Top             =   4620
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   2130
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   2790
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทที่รับกำจัด"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานะ"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Inv, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   1890
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   2550
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   7110
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   7410
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   7710
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   180
         Top             =   8010
         Width           =   195
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   7140
         Width           =   570
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รับวางบิล"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   450
         TabIndex        =   11
         Top             =   7440
         Width           =   630
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   450
         TabIndex        =   10
         Top             =   7740
         Width           =   780
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   9
         Top             =   8040
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   2460
         Y1              =   3240
         Y2              =   3240
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Inv No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อวางบิล"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "บริษัท"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ยอดก่อน VAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "VAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ยอดรวม"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่ Inv"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "วันรับชำระ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "สถานะ"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   8220
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListInv.frx":0004
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListInv.frx":059E
            Key             =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrD
    Dim tSql As String
    Dim RS As New ADODB.Recordset
    Dim tmpList As ListItem
    Dim x As Long, y As Long
    Dim sqlStr As String, strCond As String, StatusDesc As String
    Dim tmpArr, curColor
    tSql = "SELECT * FROM vw_BillHeaderSearch" & vbCrLf
    tSql = tSql & "WHERE IssuedDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "'" & vbCrLf
    If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
        tSql = tSql & " And (BillingNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or BillName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
    End If
    
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            tSql = tSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
        End If
    Else
        If cboCompany.ListIndex >= 0 Then
            tSql = tSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
        End If
    End If
    If cboStatus.ListIndex = 1 Then
        tSql = tSql & " And isCustAccepted = 'N' And isCancel <> 'Y'" & vbCrLf
    ElseIf cboStatus.ListIndex = 2 Then
        tSql = tSql & " And isCustAccepted = 'Y' And isCancel <> 'Y'" & vbCrLf
    ElseIf cboStatus.ListIndex = 3 Then
          tSql = tSql & "  And isCancel = 'Y'" & vbCrLf
    End If
    
    
        
     tSql = tSql & "ORDER BY IssuedDate, BillingNo"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    
    If RS Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    lvResult.ListItems.Clear
    Do While Not RS.EOF
        Set tmpList = lvResult.ListItems.Add(, "'" & RS!BillingHeaderID & "'", "" & RS!BillingNo)
        If Trim("" & RS!isCustAccepted) <> "Y" And Trim("" & RS!isCancel) <> "Y" Then
            curColor = lbOpen.ForeColor
            StatusDesc = lbOpen.Caption
        ElseIf Trim("" & RS!isCustAccepted) = "Y" And Trim("" & RS!isCancel) <> "Y" Then
            curColor = lbApproved.ForeColor
            StatusDesc = lbApproved.Caption
        ElseIf Trim("" & RS!isCancel) = "Y" Then
            curColor = lbCancel.ForeColor
            StatusDesc = lbCancel.Caption
        End If
        tmpList.ForeColor = curColor
        tmpList.SubItems(1) = Trim("" & RS!CompanyName)
        tmpList.SubItems(2) = Trim("" & RS!BillName)
        tmpList.SubItems(3) = Trim("" & RS!CompanyCode)
        tmpList.SubItems(4) = Format(Trim("" & RS!AmtB4VAT), "#,##0.00")
        tmpList.SubItems(5) = Format(Trim("" & RS!VATAmt), "#,##0.00")
        tmpList.SubItems(6) = Format(Trim("" & RS!TotalBillAmt), "#,##0.00")
        tmpList.SubItems(7) = FormatYMD_to_DMY(Trim("" & RS!IssuedDate), "/", "/")
        tmpList.SubItems(8) = FormatYMD_to_DMY(Trim("" & RS!PayDueDate), "/", "/")
        tmpList.SubItems(9) = StatusDesc
        For y = 1 To tmpList.ListSubItems.Count
            tmpList.ListSubItems(y).ForeColor = curColor
        Next
        Set tmpList = Nothing
        RS.MoveNext
    Loop
    lvResult.ColumnHeaders(1).Icon = "down"
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(RS.RecordCount, "#,##0") & " รายการ"
    RS.Close
    Set RS = Nothing
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    GetMsg "", vbCritical
    Exit Sub
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    PopulateCompany cboCompany, cboIndex
    With cboStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "งานใหม่"
        .AddItem "รับวางบิลแล้ว"
        .AddItem "ยกเลิก"
        .ListIndex = 0
    End With
    lvResult.ColumnHeaders(1).Icon = "down"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With lvResult
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 200
        End If
        .Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Set tmpCol = tmpS.BillingHeaderSearch("BillingHeaderID=" & lvResult.SelectedItem.Key)
    If tmpCol.Count > 0 Then
        Load frmInvoice
        frmInvoice.SetDataBillHeader tmpCol(1)
        frmInvoice.Show 1
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ClearHeaderIcons(ColumnHeader.Index)
    Select Case ColumnHeader.Index
        Case 7, 8
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortDate)
            End Select
        Case 10
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortNumeric)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortNumeric)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortNumeric)
            End Select
        Case Else
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortAlpha)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortAlpha)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortAlpha)
            End Select
    End Select
End Sub

Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim i As Integer
    For i = 1 To lvResult.ColumnHeaders.Count
        If lvResult.ColumnHeaders(i).Index <> CurrentHeader Then
            lvResult.ColumnHeaders(i).Icon = Empty
        End If
    Next
End Sub

Private Sub txtKey_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Trim(txtKey.Text) <> "" Then cmdSearch_Click
End Sub
