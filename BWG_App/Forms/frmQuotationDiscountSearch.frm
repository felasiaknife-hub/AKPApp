VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#3.0#0"; "DateCtl.ocx"
Begin VB.Form frmQuotationDiscountSearch 
   Caption         =   "สวนลดใบเสนอราคา"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   120
      Picture         =   "frmQuotationDiscountSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   915
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2055
      TabIndex        =   1
      Text            =   "*"
      Top             =   1380
      Width           =   4305
   End
   Begin DateCtl.Date dtFrom 
      Height          =   315
      Left            =   2055
      TabIndex        =   0
      Top             =   960
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin DateCtl.Date dtTo 
      Height          =   315
      Left            =   4575
      TabIndex        =   2
      Top             =   960
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3345
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   5900
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
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ใบสวนลด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ลูกค้า"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ผลการค้นหา"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   9015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   5805
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   2
      X1              =   0
      X2              =   5715
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frmQuotationDiscountSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสนอราคา, ชื่อลูกค้า, รหัสลูกค้า"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2145
      TabIndex        =   7
      Top             =   1740
      Width           =   2685
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาใบเสนอราคา"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   135
      TabIndex        =   6
      Top             =   75
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   45
      X2              =   5850
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   45
      X2              =   5760
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   615
      TabIndex        =   5
      Top             =   1410
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ในช่วงวันที่"
      Height          =   195
      Index           =   0
      Left            =   1155
      TabIndex        =   4
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   1
      Left            =   3975
      TabIndex        =   3
      Top             =   1020
      Width           =   495
   End
End
Attribute VB_Name = "frmQuotationDiscountSearch"
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

Private Sub cmdOpen_Click()
   If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลใบเสนอราคาที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกข้อมูลใบเสนอราคาที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataQuotationDiscount colSearchResult_(lvResult.SelectedItem.Key)
'    Unload Me
    Me.Hide
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result%, y%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGQTDiscountHeader
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
    strCond = "(CustomerName Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%' Or QuoDisNo Like '%" & criStr & "%')"
    Set colSearchResult_ = tmpS.QuotationDiscountHeaderSearch(strCond)
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.QuoDisNo)
        tmpList.SubItems(1) = Trim(tmpItem.QuoDisNo)
        tmpList.SubItems(2) = Trim(tmpItem.CustomerName)
        tmpList.SubItems(3) = Trim(tmpItem.CreateDate)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpS = Nothing
End Sub

Private Sub Form_Activate()
    
        cmdSearch_Click
'        firstActive = False
'        If lvResult.ListItems.Count = 1 Then
'            callerForm_.SetDataCustomer colSearchResult_(1)
'            Unload Me
'        End If
End Sub
Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    dtFrom.ValueYMD = Left(TodayDate, 8) & "01"
    dtTo.ValueYMD = TodayDate
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
    cmdOpen_Click
End Sub
