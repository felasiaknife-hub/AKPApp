VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaleDistrictSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F115"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   45
      Picture         =   "frSaleDistrictSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5235
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6045
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5235
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5235
      Width           =   915
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Text            =   "*"
      Top             =   900
      Width           =   4305
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3375
      Left            =   0
      TabIndex        =   4
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
         Text            =   "เขต"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "โซน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "พนักงานขาย"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อเขต, โซน, พนักงานขายที่รับผิดชอบ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   8
      Top             =   1260
      Width           =   2640
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
      Left            =   -90
      TabIndex        =   7
      Top             =   1605
      Width           =   7125
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -30
      Top             =   1515
      Width           =   7125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   945
      Width           =   1365
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาเขตพื้นที่การขาย"
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
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   2205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   -15
      X2              =   5790
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   -15
      X2              =   5700
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frSaleDistrictSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmSaleDistrictSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean

Public Sub getCriteria(criStr As String)
    firstActive = True
    txtCriteria.Text = Trim(criStr)
End Sub

Private Sub Form_Activate()
    If firstActive Then
        cmdSearch_Click
        firstActive = False
        If lvResult.ListItems.Count = 1 Then
            callerForm_.SetDataSaleDistrict colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลเขตการขายที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกเขตการขายที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataSaleDistrict colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim x%, result%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGSaleDistrict
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
    strCond = "(StaffFName Like '%" & criStr & "%' Or StaffLName Like '%" & criStr & "%' Or ZoneName Like '%" & criStr & "%' Or SDName Like '%" & criStr & "%')"
    
    Set colSearchResult_ = tmpS.SaleDistrictSearch(strCond & " AND isActive = 'Y'")
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.DistrictName)
        tmpList.SubItems(1) = Trim(tmpItem.ZoneName)
        tmpList.SubItems(2) = Trim(tmpItem.SaleStaffName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpS = Nothing

End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOpen_Click
End Sub
