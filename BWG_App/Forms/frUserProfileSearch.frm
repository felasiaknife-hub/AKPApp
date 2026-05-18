VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmUserProfileSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F144"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUserGroup 
      Height          =   315
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1950
      Width           =   4275
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   4275
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   45
      Picture         =   "frUserProfileSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6315
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6315
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6315
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
      Height          =   3525
      Left            =   0
      TabIndex        =   6
      Top             =   2730
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   6218
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
         Text            =   "รหัส"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อ-นามกุล"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Username"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กลุ่มผู้ใช้ระบบ"
      Height          =   195
      Index           =   1
      Left            =   615
      TabIndex        =   12
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บริษัทที่สังกัด"
      Height          =   195
      Index           =   0
      Left            =   645
      TabIndex        =   11
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัส, ชื่อ-นามสกุล,ชื่อในการเข้าระบบ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   10
      Top             =   1260
      Width           =   2535
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
      TabIndex        =   9
      Top             =   2505
      Width           =   7605
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -30
      Top             =   2415
      Width           =   7605
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   8
      Top             =   915
      Width           =   1365
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาข้อมูลผู้ใช้ระบบ"
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
      TabIndex        =   7
      Top             =   0
      Width           =   2115
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
      Picture         =   "frUserProfileSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmUserProfileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean
Dim colCompany_ As Collection
Dim colUserGroup_ As Collection

Public Sub getCriteria(criStr As String)
    firstActive = True
    txtCriteria.Text = Trim(criStr)
End Sub

Private Sub cboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub cboUserGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub Form_Activate()
    If firstActive Then
        cmdSearch_Click
        firstActive = False
        If lvResult.ListItems.Count = 1 Then
            callerForm_.setCurrentUser colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูล Supplier ที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือก Supplier ที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.setCurrentUser colSearchResult_(lvResult.SelectedItem.Key)
'    Unload Me
    Me.Hide
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGCompanyStaff
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
    strCond = "(StaffCode Like '%" & criStr & "%' Or StaffFName Like '%" & criStr & "%' Or StaffLName Like '%" & criStr & "%' Or LoginName Like '%" & criStr & "%')"
    If cboCompany.ListIndex > 0 Then
        strCond = strCond & " And StaffID in (Select StaffID from tbStaffInCompany WHERE CompanyID='" & colCompany_(cboCompany.ListIndex).ID & "')"
    End If
    If cboUserGroup.ListIndex > 0 Then
        strCond = strCond & " And StaffID in (Select StaffID from tbUserInGroup WHERE GroupID='" & ColUserGroup(cboUserGroup.ListIndex).ID & "')"
    End If
    Set colSearchResult_ = tmpS.UserSearch(strCond)
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.StaffCode)
        tmpList.SubItems(1) = Trim(tmpItem.StaffFName & " " & tmpItem.StaffLName)
        tmpList.SubItems(2) = Trim(tmpItem.LoginName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        If firstActive = True Then
            result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลผู้ใช้ระบบใหม่หรือไม่", vbQuestion + vbYesNo)
            If result = vbYes Then
                frmUserProfile.cmdSearch.Enabled = False
                frmUserProfile.Show 1
                Set colSearchResult_ = tmpS.UserSearch(strCond)
                For x = 1 To colSearchResult_.Count
                    Set tmpItem = colSearchResult_(x)
                    Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.StaffCode)
                    tmpList.SubItems(1) = Trim(tmpItem.StaffFName & " " & tmpItem.StaffLName)
                    tmpList.SubItems(2) = Trim(tmpItem.LoginName)
                    Set tmpItem = Nothing
                    Set tmpList = Nothing
                Next
            End If
        Else
            MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
        End If
    End If
    Set tmpS = Nothing
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub Form_Load()
Dim x%
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    
    PopulateCompany
    PopulateUserGroup
    cboUserGroup.Clear
    For x = 1 To ColUserGroup.Count
        cboUserGroup.AddItem ColUserGroup(x).Name
    Next
    cboUserGroup.AddItem "== ทุกกลุ่มผู้ใช้ ==", 0
    cboUserGroup.ListIndex = 0
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

Private Sub PopulateCompany()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colCompany_ = tmpS.CompanySearch
    Set tmpS = Nothing
    cboCompany.Clear
    For x = 1 To colCompany_.Count
        cboCompany.AddItem colCompany_(x).CompanyNameTH
    Next
    cboCompany.AddItem "== ทุกบริษัท ==", 0
    cboCompany.ListIndex = 0
End Sub

