VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerExtComm 
   Caption         =   "รายชื่อลูกค้าที่ต้องจ่ายค่าคอมนอก"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   2655
      Begin VB.CommandButton cmdPrint 
         Caption         =   "พิมพ์ข้อมูล"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2940
         Width           =   2115
      End
      Begin VB.TextBox txtCriteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Text            =   "*"
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Text            =   "*"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   2115
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3960
         Width           =   2115
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   120
         Top             =   7380
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
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
         TabIndex        =   3
         Top             =   300
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3495
      Left            =   2640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัสลูกค้า"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ผู้แทนขาย"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "จังหวัด"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmCustomerExtComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean

Dim curStaff_ As BWGCompanyStaff

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim tmpSql As String, iCount As Long
Dim criStr As String, strCond As String
    criStr = Replace(Replace(Trim(txtCriteria.Text), "*", "%"), "'", "")
    strCond = "(CustomerName Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%')"
    
    tmpSql = "Select  DISTINCT CustomerID, CustomerCode, CustomerName, ProvinceName, SaleStaffFName, SaleStaffLName  from vw_CustomerSearch_ExtComm Where (" & strCond & ") Order By CustomerName"

    With rpt1
        .Reset
        .WindowTitle = "รายลูกค้าที่ต้องจ่ายค่าคอมฯ ภายนอก"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\CustomerComEx.rpt"
        .SqlQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With

End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim x%, Result
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGCustomer
    criStr = Replace(Replace(Trim(txtCriteria.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtCriteria.SetFocus
        Exit Sub
    End If
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    strCond = "(CustomerName Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%')"
    If callerForm_.Name = "frCustomer" Then
        'If frCustomer.CallLocation = "P" Then strCond = "(" & strCond & " And CustTypeFlag='C')"
    End If
    Set colSearchResult_ = tmpS.CustomerSearchExtComm(strCond)
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.CustomerCode)
        tmpList.SubItems(1) = Trim(tmpItem.CustomerName)
        tmpList.SubItems(2) = Trim(tmpItem.UnderSaleName)
        tmpList.SubItems(3) = Trim(tmpItem.ProvinceName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation, "Result"
    End If
    Set tmpS = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lvResult.Width = Me.Width - Frame1.Width - 100
    lvResult.Height = Me.Height - 520
    Frame1.Height = Me.Height - 410
    
    If lvResult.Width > (lvResult.ColumnHeaders(1).Width + lvResult.ColumnHeaders(3).Width + lvResult.ColumnHeaders(4).Width + 2500) Then
        lvResult.ColumnHeaders(2).Width = lvResult.Width - (lvResult.ColumnHeaders(1).Width + lvResult.ColumnHeaders(3).Width + lvResult.ColumnHeaders(4).Width) - 500
    End If
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
    lvResult.Sorted = True
End Sub

Private Sub lvResult_DblClick()
    frmExtCom2.txtCustomer.Text = lvResult.SelectedItem.Text
    frmExtCom2.SetDataCustomer colSearchResult_(lvResult.SelectedItem.Key)
    Load frmExtCom2
    frmExtCom2.Show
    frmExtCom2.ZOrder
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If txtSale.Text = "" Then Exit Sub
    If txtSale.Text = "*" Then Exit Sub
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtSale.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub
