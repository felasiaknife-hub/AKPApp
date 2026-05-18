VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuotationSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F112"
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
   Begin VB.ComboBox cboCustType 
      Height          =   315
      Left            =   5640
      TabIndex        =   14
      Top             =   1400
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
      Width           =   2505
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
      Picture         =   "frQuotationSearch.frx":0000
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
         Text            =   "เลขที่ใบเสนอราคา"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ลูกค้า"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "สถานะ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ประเภท"
         Object.Width           =   2117
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
      Caption         =   "ประเภทลูกค้า"
      Height          =   195
      Index           =   3
      Left            =   4635
      TabIndex        =   13
      Top             =   1440
      Width           =   930
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
      Left            =   120
      TabIndex        =   9
      Top             =   75
      Width           =   1620
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
      Caption         =   "เลขที่ใบเสนอราคา, ชื่อลูกค้า, รหัสลูกค้า"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2130
      TabIndex        =   7
      Top             =   1740
      Width           =   2685
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frQuotationSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmQuotationSearch"
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
'    Unload Me
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
    callerForm_.SetDataQuotation colSearchResult_(lvResult.SelectedItem.Key)
'    Unload Me
    Me.Hide
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
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(F) RDF Waste"
        .AddItem "(L) (p)"
        .AddItem "(T) (k)"
        .AddItem "(G) (e)"
        .AddItem "(R)"
        .ListIndex = 0
    End With
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
On Error GoTo ErrD
Dim criStr As String
Dim strCond As String
Dim x%, result%, Y%
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
    strCond = "(QuotationNo Like '%" & criStr & "%') OR (CustomerCode Like '%" & criStr & "%') OR (CompanyName Like '%" & criStr & "%'))"
    If firstActive = False Then
        strCond = strCond & " AND ((IssueDate>='" & dtFrom.ValueYMD & "' AND IssueDate<='" & dtTo.ValueYMD & "')"
    End If
    If frmQuotationDiscout.isQuotationDiscount = True Then
        strCond = strCond & " And (isQuoDisount = 'N' Or isQuoDisount is null )  "
    End If
    If cboCustType.ListIndex = 1 Then
        strCond = strCond & " AND CustTypeFlag in ('C','S') "
    ElseIf cboCustType.ListIndex = 2 Then
        strCond = strCond & " AND CustTypeFlag = 'F' "
   ElseIf cboCustType.ListIndex = 3 Then
        strCond = strCond & " AND CustTypeFlag in ('L','P') "
    ElseIf cboCustType.ListIndex = 4 Then
        strCond = strCond & " AND CustTypeFlag in ('T','K') "
    ElseIf cboCustType.ListIndex = 5 Then
        strCond = strCond & " AND CustTypeFlag in ('G','E') "
    ElseIf cboCustType.ListIndex = 6 Then
        strCond = strCond & " AND CustTypeFlag in ('R') "
    End If
    strCond = SelectWithCompanyID(strCond, True)
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
            For Y = 1 To 3
                tmpList.ListSubItems(Y).Bold = True
                tmpList.ListSubItems(Y).ForeColor = vbRed
            Next
        ElseIf tmpItem.isclosed Then
            tmpList.SubItems(3) = "Closed"
            tmpList.ForeColor = vbRed
            tmpList.Bold = True
            For Y = 1 To 3
                tmpList.ListSubItems(Y).Bold = True
                tmpList.ListSubItems(Y).ForeColor = vbRed
            Next
        ElseIf tmpItem.isApproved Then
            tmpList.SubItems(3) = "อนุมัติแล้ว"
            tmpList.ForeColor = &H8000&
            tmpList.Bold = True
            For Y = 1 To 3
                tmpList.ListSubItems(Y).Bold = True
                tmpList.ListSubItems(Y).ForeColor = &H8000&
            Next
        Else
            tmpList.SubItems(3) = "ยังไม่อนุมัติ"
        End If
        
        If tmpItem.isJob = "N" Then
            tmpList.SubItems(4) = "งาน Routine"
        ElseIf tmpItem.isJob = "Y" Then
            tmpList.SubItems(4) = "งาน Jobs"
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
    Exit Sub
    
ErrD:
    MsgBox err.Description, vbCritical
    Exit Sub
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOpen_Click
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub
