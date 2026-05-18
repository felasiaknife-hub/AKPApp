VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmListMenifestBill 
   Caption         =   "รายการ Menifest"
   ClientHeight    =   9000
   ClientLeft      =   1920
   ClientTop       =   -15
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      TabIndex        =   8
      Top             =   -120
      Width           =   2655
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3300
         Width           =   2205
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   2040
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtFrom 
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   1320
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cboConditon 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   2700
         Width           =   2235
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "พิมพ์ข้อมูล"
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4380
         Width           =   2235
      End
      Begin VB.TextBox txtCustomer 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5100
         Width           =   2235
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   2235
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   300
         Top             =   8400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label lblNew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   570
         TabIndex        =   18
         Top             =   5910
         Width           =   570
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   300
         Top             =   5880
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   300
         Top             =   6180
         Width           =   195
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วางบิลแล้ว"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   570
         TabIndex        =   17
         Top             =   6210
         Width           =   735
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   300
         Top             =   6480
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   300
         Top             =   6780
         Width           =   195
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   570
         TabIndex        =   16
         Top             =   6510
         Width           =   780
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   570
         TabIndex        =   15
         Top             =   6810
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "บริษัท"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   3060
         Width           =   555
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "สถานะ"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "วันที่ :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   1740
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า , ผู้แทนขาย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   1560
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3495
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ใบคุม"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลข Menifest"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เลข Invoice"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่ Invoice"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ลูกค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "น้ำหนัก"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ค่าบริการ"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "เลข RE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "วันที่ ออก RE"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "พนังงานขาย"
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "frmListMenifestBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim tmpSql As String, iCount As Long
Dim criStr As String, strCond As String
    criStr = Replace(Replace(Trim(txtCustomer.Text), "*", "%"), "'", "")
    If txtCustomer.Text = Trim("") Or txtCustomer.Text = Trim("*") Then
        strCond = " WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "'"
    Else
        strCond = "(CustomerName Like '%" & criStr & "%' Or SaleFName Like '%" & criStr & "%') And WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "'"
    End If
    If cboConditon.ListIndex = 1 Then
        strCond = strCond & " AND ReNo IS NULL"
    ElseIf cboConditon.ListIndex = 2 Then
        strCond = strCond & " AND ReNo IS NOT NULL"
    End If
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
        End If
    Else
        If cboCompany.ListIndex >= 0 Then
            strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
        End If
    End If
    tmpSql = "Select  DISTINCT MenifestID, DocNo, MenifestNo, WorkDate, CustomerName,CompanyNameTH, SaleFName, SaleLName, RENo, REDate, InCollectFee, ActualQty from vw_Me_Bill_RE WHERE (" & strCond & ") ORDER BY WorkDate"
    With rpt1
        .Reset
        .WindowTitle = "รายการใบกำกับใบขนส่ง (Menifest No.)"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\ListMenifest.rpt"
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim x%, Y%, result
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGMenifestBillRE
Dim varColor
    criStr = Replace(Replace(Trim(txtCustomer.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    If txtCustomer.Text = Trim("") Or txtCustomer.Text = Trim("*") Then
        strCond = " WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "'"
    Else
        strCond = "(CustomerName Like '%" & criStr & "%' Or SaleFName Like '%" & criStr & "%') And WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "'"
    End If
    If cboConditon.ListIndex = 1 Then
        strCond = strCond & " AND ReNo IS NULL"
    ElseIf cboConditon.ListIndex = 2 Then
        strCond = strCond & " AND ReNo IS NOT NULL"
    End If
    
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
        End If
    Else
        If cboCompany.ListIndex >= 0 Then
            strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
        End If
    End If
    
    Set colSearchResult_ = tmpS.MenifestSearchRE(strCond)
    If colSearchResult_.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation, "Result": Exit Sub
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", x)
        tmpList.SubItems(1) = tmpItem.TimeTableNo
        tmpList.SubItems(2) = Trim(tmpItem.docNo)
        tmpList.SubItems(3) = Trim(tmpItem.WorkDate)
        tmpList.SubItems(4) = Trim(tmpItem.BillingNo)
        tmpList.SubItems(5) = Trim(tmpItem.IssuedDate)
        tmpList.SubItems(6) = Trim(tmpItem.CustomerName)
        tmpList.SubItems(7) = Trim(tmpItem.ActualQty)
        tmpList.SubItems(8) = Trim(tmpItem.InCollectFee)
        tmpList.SubItems(9) = Trim(tmpItem.RENo)
        tmpList.SubItems(10) = Trim(tmpItem.REDate)
        tmpList.SubItems(11) = Trim(tmpItem.SaleFName) & " " & Trim(tmpItem.SaleLName)
        Select Case UCase(tmpItem.MStatus)
            Case "CC"
                varColor = lbCancel.ForeColor
            Case "CL"
                varColor = lbClosed.ForeColor
            Case "IV"
                varColor = lblInv.ForeColor
            Case Else
                varColor = lblNew.ForeColor
        End Select
        tmpList.ForeColor = varColor
        For Y = 1 To tmpList.ListSubItems.Count
            tmpList.ListSubItems(Y).ForeColor = varColor
        Next
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
'    If lvResult.ListItems.Count = 0 Then
'        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation, "Result"
'    End If
    Set tmpS = Nothing
End Sub

Private Sub Form_Load()
    cboConditon.AddItem "เลือกทั้งหมด"
    cboConditon.AddItem "รายการที่ยังไม่ได้เก็บเงิน"
    cboConditon.AddItem "รายการที่เก็บเงินได้"
    cboConditon.ListIndex = 1
    
    PopulateCompany cboCompany, cboIndex
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lvResult.Width = Me.Width - Frame1.Width - 100
    lvResult.Height = Me.Height - 520
    Frame1.Height = Me.Height - 410
    
'    If lvResult.Width > (lvResult.ColumnHeaders(2).Width + lvResult.ColumnHeaders(3).Width + lvResult.ColumnHeaders(4).Width + 2500) Then
'        lvResult.ColumnHeaders(1).Width = lvResult.Width - (lvResult.ColumnHeaders(2).Width + lvResult.ColumnHeaders(3).Width + lvResult.ColumnHeaders(4).Width) - 500
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim i As Long
    lvResult.Sorted = True
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
    For i = 1 To lvResult.ListItems.Count
        lvResult.ListItems(i).Text = i
    Next
End Sub



