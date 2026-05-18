VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCarQuotationSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   0
      Picture         =   "frmCarQuotationSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   915
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Text            =   "*"
      Top             =   960
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   1830
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   5953
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ประเภทกลุ่มหลัก"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ประเภทกลุ่มย่อย"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Rate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ผลการค้นหา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1605
      Width           =   7125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทกลุ่มหลัก, ประเภทกลุ่มย่อย"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -30
      Top             =   1560
      Width           =   7125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   2
      Top             =   1005
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5715
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   5805
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาประเภทรถ"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frmCarQuotationSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmCarQuotationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuotationID_ As String
Dim CarNo_  As String
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Private Sub cmdClose_Click()
    Unload Me
End Sub
Public Sub SetCallerForm(frm As Form, Optional QuotationID As String, Optional carNo As String)
    Set callerForm_ = frm
    QuotationID_ = QuotationID
    CarNo_ = carNo
   cmdSearch_Click
End Sub

Private Sub cmdOpen_Click()
 If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลพนักงานที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกพนักงานที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataTruckSubTypeRate colSearchResult_(lvResult.SelectedItem.Key), CarNo_
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGQuotationTruckRate
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

    Set colSearchResult_ = tmpS.QuotationTrucSubTypekRateSearch("QuotationID in (" & QuotationID_ & ")")
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.id & "'", tmpItem.TypeDesc)
        tmpList.SubItems(1) = Trim(tmpItem.SubTypeDesc)
        tmpList.SubItems(2) = Trim(tmpItem.Rate)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpS = Nothing
End Sub

Private Sub lvResult_DblClick()
cmdOpen_Click
End Sub
