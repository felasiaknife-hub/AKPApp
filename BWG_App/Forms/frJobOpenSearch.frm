VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobOpenSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F043"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frJobOpenSearch.frx":0000
      Left            =   6420
      List            =   "frJobOpenSearch.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
   Begin DateCtl.Date dtFrom 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "*"
      Top             =   1380
      Width           =   4305
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5910
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5910
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   30
      Picture         =   "frJobOpenSearch.frx":002B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5910
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3465
      Left            =   0
      TabIndex        =   3
      Top             =   2370
      Width           =   9105
      _ExtentX        =   16060
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบสั่งขน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่สั่งงาน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ขน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ลูกค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ประเภท job"
         Object.Width           =   2540
      EndProperty
   End
   Begin DateCtl.Date dtTo 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   960
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   12
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ในช่วงวันที่"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   11
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   600
      TabIndex        =   10
      Top             =   1410
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
      Caption         =   "ค้นหาใบสั่งขนกากของเสีย"
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
      Left            =   120
      TabIndex        =   9
      Top             =   75
      Width           =   2265
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   0
      Top             =   2070
      Width           =   9165
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
      TabIndex        =   8
      Top             =   2130
      Width           =   9285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งขน, ชื่อลูกค้า, รหัสลูกค้า"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2130
      TabIndex        =   7
      Top             =   1740
      Width           =   2340
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frJobOpenSearch.frx":27CD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmJobOpenSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean

Public JobType As String

Public Sub getCriteria(criStr As String)
    firstActive = True
    txtCriteria.Text = Trim(criStr)
End Sub

Private Sub cboDateType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdClose_Click()
    Me.Hide
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
    If isExist(CurrentUser.WorkCompanies, "'" & colSearchResult_(lvResult.SelectedItem.Key).CompanyID & "'") = False Then
        MsgBox "ท่านไม่สามารถเปิด Order ของบริษัทอื่นได้", vbExclamation
        Exit Sub
    Else
        callerForm_.SetDataJobOpen colSearchResult_(lvResult.SelectedItem.Key)
        'Unload Me
        Me.Hide
    End If
End Sub

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
    If firstActive Then
        cmdSearch_Click
        firstActive = False
        If lvResult.ListItems.Count = 1 Then
            callerForm_.SetDataJobOpen colSearchResult_(1)
            'Unload Me
            Me.Hide
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Combo1.ListIndex = 0
    dtFrom.ValueYMD = Left(TodayDate, 8) & "01"
    dtTo.ValueYMD = TodayDate
End Sub

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
    cmdOpen_Click
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim x%, result%, y%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem
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
    strCond = "(JobNo Like '%" & criStr & "%') OR (CustomerCode Like '%" & criStr & "%') OR (CompanyName Like '%" & criStr & "%')) AND (ISNULL(JobNo, '') <> '' "
    If firstActive = False Then
        If Combo1.ListIndex = 0 Then
            strCond = strCond & " AND (WorkDate>='" & dtFrom.ValueYMD & "' AND WorkDate<='" & dtTo.ValueYMD & "')"
        Else
            strCond = strCond & " AND (IssuedDate>='" & dtFrom.ValueYMD & "' AND IssuedDate<='" & dtTo.ValueYMD & "')"
        End If
    End If
    
    strCond = SelectWithCompanyID(strCond, True)
    If JobType = "" Then
        JobType = "JobDataHeaderSearch"
    End If
    
    If JobType = "JobDataHeaderSearch" Then
        Set colSearchResult_ = tmpS.JobDataHeaderSearch(strCond)
        Set tmpItem = New BWGJobData
    ElseIf JobType = "JobDataHeaderBoxSearch" Then
        Set colSearchResult_ = tmpS.JobDataHeaderBoxSearch(strCond)
        Set tmpItem = New BWGJobDataBox
    End If
    
    Set tmpS = Nothing
    
    
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.JobNo)
        tmpList.SubItems(1) = FormatYMD_to_DMY(Trim(tmpItem.IssuedDate), "/", "/")
        tmpList.SubItems(2) = FormatYMD_to_DMY(Trim(tmpItem.WorkDate), "/", "/")
        tmpList.SubItems(3) = Trim(tmpItem.CustomerName)
        If Trim(tmpItem.isJob) = "N" Then
            tmpList.SubItems(4) = "Routine"
        ElseIf Trim(tmpItem.isJob) = "Y" Then
            tmpList.SubItems(4) = "Jobs"
        End If
        If tmpItem.isCancel Then
            tmpList.ForeColor = vbRed
            tmpList.ListSubItems(1).ForeColor = vbRed
            tmpList.ListSubItems(2).ForeColor = vbRed
            tmpList.ListSubItems(3).ForeColor = vbRed
            tmpList.ListSubItems(4).ForeColor = vbRed
        ElseIf tmpItem.isCompleted Then
            tmpList.ForeColor = vbGreen
            tmpList.ListSubItems(1).ForeColor = vbGreen
            tmpList.ListSubItems(2).ForeColor = vbGreen
            tmpList.ListSubItems(3).ForeColor = vbGreen
            tmpList.ListSubItems(4).ForeColor = vbGreen
        End If
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
'    If lvResult.ListItems.Count = 0 Then
'        result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลกลุ่มหลักสินค้าใหม่หรือไม่", vbQuestion + vbYesNo)
'        If result = vbYes Then
'            frProductSubGroup.cmdSearch.Enabled = False
'            frProductSubGroup.Show 1
'            Set tmpS = New BWGSearchManager
'            Set colSearchResult_ = tmpS.POHeaderSearch(strCond)
'            For x = 1 To colSearchResult_.Count
'                Set tmpItem = colSearchResult_(x)
'                Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.PONo)
'                tmpList.SubItems(1) = Trim(tmpItem.POStatus)
'                Set tmpItem = Nothing
'                Set tmpList = Nothing
'            Next
'        End If
'    End If
    Set tmpS = Nothing
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOpen_Click
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub
