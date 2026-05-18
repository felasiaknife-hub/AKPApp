VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frCustomerSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Text            =   "*"
      Top             =   990
      Width           =   4305
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5550
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5550
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   60
      Picture         =   "frCustomerSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5550
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3465
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2010
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   6112
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
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสลูกค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ผู้แทนขาย"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "จังหวัด"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   30
      X2              =   5745
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   30
      X2              =   5835
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาข้อมูลลูกค้า"
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
      Left            =   120
      TabIndex        =   7
      Top             =   45
      Width           =   1680
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   75
      Index           =   1
      Left            =   0
      Top             =   1680
      Width           =   7935
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
      TabIndex        =   6
      Top             =   1770
      Width           =   8055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ, ชื่อกลุ่มลูกค้า"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   5
      Top             =   1350
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frCustomerSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7860
   End
End
Attribute VB_Name = "frCustomerSearch"
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
            callerForm_.SetDataCustomer colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลลูกค้าที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกลูกค้าที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataCustomer colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result
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
    strCond = "(CompanyName Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%')"
    If callerForm_.Name = "frCustomer" Then
        'If frCustomer.CallLocation = "P" Then strCond = "(" & strCond & " And CustTypeFlag='C')"
    End If
    Set colSearchResult_ = tmpS.CustomerSearch(strCond)
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.CustomerName)
        tmpList.SubItems(1) = Trim(tmpItem.CustomerCode)
        tmpList.SubItems(2) = Trim(tmpItem.UnderSaleName)
        tmpList.SubItems(3) = Trim(tmpItem.ProvinceName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        If firstActive = True Then
            result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลลูกค้าใหม่หรือไม่", vbQuestion + vbYesNo)
            If result = vbYes Then
'                frCustomer.cmdSearch.Enabled = False
'                frCustomer.Show 1
'                Set colSearchResult_ = tmpS.CustomerSearch(strCond)
'                For x = 1 To colSearchResult_.Count
'                    Set tmpItem = colSearchResult_(x)
'                    Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.CustomerName)
'                    tmpList.SubItems(1) = Trim(tmpItem.CustomerCode)
'                    tmpList.SubItems(2) = Trim(tmpItem.UnderSaleName)
'                    tmpList.SubItems(3) = Trim(tmpItem.ProvinceName)
'                    Set tmpItem = Nothing
'                    Set tmpList = Nothing
'                Next
            End If
        Else
            MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
        End If
    End If
    Set tmpS = Nothing
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Public Sub setCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As BWGCustomer
Dim tmpStr As String
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New BWGCustomer
            tmpItem.ID = Replace(lvResult.SelectedItem.Key, "'", "")
            tmpStr = tmpItem.Delete
            If Trim(tmpStr) = "" Then
                lvResult.ListItems.Remove lvResult.SelectedItem.Index
                Set tmpItem = Nothing
                txtCriteria.Text = ""
                txtCriteria.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
    If KeyCode = 13 Then cmdOpen_Click

End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click

End Sub



