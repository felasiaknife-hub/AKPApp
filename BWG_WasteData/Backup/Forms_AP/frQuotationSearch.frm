VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frQuotationSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   $"frQuotationSearch.frx":0000
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin WasteDataForm.ctlDate dtTo 
      Height          =   315
      Left            =   4620
      TabIndex        =   12
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteDataForm.ctlDate dtFrom 
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "*"
      Top             =   1380
      Width           =   4305
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5910
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5910
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   30
      Picture         =   "frQuotationSearch.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5910
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3465
      Left            =   0
      TabIndex        =   1
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบเสนอราคา"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ลูกค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "สถานะ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      Caption         =   "ค้นหาใบสั่งซื้อสินค้า"
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
      TabIndex        =   7
      Top             =   75
      Width           =   1755
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
      TabIndex        =   6
      Top             =   2130
      Width           =   9285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสนอราคา, ชื่อลูกค้า, รหัสลูกค้า"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2130
      TabIndex        =   5
      Top             =   1740
      Width           =   2685
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frQuotationSearch.frx":27AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frQuotationSearch"
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

Private Sub cboDateType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

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
    callerForm_.SetDataQuotation colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
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
            callerForm_.SetDataQuotation colSearchResult_(1)
            Unload Me
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
    dtFrom.ValueYMD = Left(TodayDate, 8) & "01"
    dtTo.ValueYMD = TodayDate
End Sub

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result%, y%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGQuotationHeader
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
    strCond = "(QuotationNo Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%' Or CompanyName Like '%" & criStr & "%')"
    If firstActive = False Then
        strCond = strCond & " And (IssueDate>='" & dtFrom.ValueYMD & "' And IssueDate<='" & dtTo.ValueYMD & "')"
    End If
    Set colSearchResult_ = tmpS.QuotationHeaderSearch(strCond)
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.QuotationNo)
        tmpList.SubItems(1) = FormatYMD_to_DMY(Trim(tmpItem.IssuedDate), "/", "/")
        tmpList.SubItems(2) = Trim(tmpItem.CustomerName)
        If tmpItem.isCancel Then
            tmpList.SubItems(3) = "ยกเลิก"
            tmpList.ForeColor = vbRed
            tmpList.Bold = True
            For y = 1 To 3
                tmpList.ListSubItems(y).Bold = True
                tmpList.ListSubItems(y).ForeColor = vbRed
            Next
        ElseIf tmpItem.isClosed Then
            tmpList.SubItems(3) = "Closed"
            tmpList.ForeColor = vbRed
            tmpList.Bold = True
            For y = 1 To 3
                tmpList.ListSubItems(y).Bold = True
                tmpList.ListSubItems(y).ForeColor = vbRed
            Next
        ElseIf tmpItem.IsApproved Then
            tmpList.SubItems(3) = "อนุมัติแล้ว"
            tmpList.ForeColor = &H8000&
            tmpList.Bold = True
            For y = 1 To 3
                tmpList.ListSubItems(y).Bold = True
                tmpList.ListSubItems(y).ForeColor = &H8000&
            Next
        Else
            tmpList.SubItems(3) = "ยังไม่อนุมัติ"
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
