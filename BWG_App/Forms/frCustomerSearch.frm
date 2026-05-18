VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F026"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCustType 
      Height          =   315
      Left            =   5040
      TabIndex        =   13
      Top             =   1850
      Width           =   1815
   End
   Begin VB.TextBox txtSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Text            =   "*"
      Top             =   1380
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "แสดงเฉพาะลูกค้าที่ยกเลิก"
      Height          =   195
      Left            =   1620
      TabIndex        =   10
      Top             =   1860
      Width           =   2175
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Text            =   "*"
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   675
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6120
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   675
      Left            =   60
      Picture         =   "frCustomerSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3585
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2490
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   6324
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
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสลูกค้า"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขทะเบียนโรงงาน"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ผู้แทนขาย"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "จังหวัด"
         Object.Width           =   2364
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   12
      Top             =   1855
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้แทนขาย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   780
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   1395
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
      TabIndex        =   8
      Top             =   45
      Width           =   1680
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   0
      Top             =   1800
      Width           =   11440
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
      Height          =   285
      Index           =   2
      Left            =   -90
      TabIndex        =   7
      Top             =   2220
      Width           =   10575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ, รหัสลูกค้า,เลขทะเบียนโรงงาน"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   6060
      TabIndex        =   6
      Top             =   1020
      Width           =   2340
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frCustomerSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11440
   End
End
Attribute VB_Name = "frmCustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean
Dim strOptional As String

Public CStatus As String
Dim curStaff_ As BWGCompanyStaff
Dim ShowSubContact_ As String

Public Sub getCriteria(criStr As String)
    firstActive = True
    txtCriteria.Text = Trim(criStr)
End Sub

Public Sub GetCriteriaOptonal(Optional ByVal strOpt As String = "")
    strOptional = strOpt
End Sub

Public Property Get ShowSubContact() As Boolean
    If ShowSubContact_ = "Y" Then
        ShowSubContact = True
    Else
        ShowSubContact = False
    End If
End Property

Public Property Let ShowSubContact(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(ShowSubContact_) <> Trim(tmpStr) Then
        ShowSubContact_ = Trim(tmpStr)
    End If
End Property

Private Sub Form_Activate()
    If firstActive Then
        cmdSearch_Click
        firstActive = False
        If lvResult.ListItems.Count = 1 Then
            callerForm_.SetDataCustomer colSearchResult_(1)
            Unload Me
        End If
    End If
    ShowSubContact = True
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
On Error Resume Next
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
    strCond = "(CompanyName Like '%" & criStr & "%' OR CustomerCode Like '%" & criStr & "%' or RegisterNo Like '%" & criStr & "%' or [14digit] Like '%" & criStr & "%')"
    If callerForm_.Name = "frmCustomer" Then
        If callerForm_.CallLocation = "P" Then strCond = "(" & strCond & " AND CustTypeFlag='C')"
    End If
    
    If Not curStaff_ Is Nothing Then
        If curStaff_.ID <> "" Then
            strCond = strCond & " AND UnderSaleStaffID = '" & curStaff_.ID & "'"
        End If
    End If
    
    If CStatus = "" Then
        CStatus = "N"
    End If
        
    If Check1.Value = 1 Then
'        strCond = strCond & " AND (Status = '" & Trim(CStatus) & "'"
'        strCond = strCond & " OR Status = 'C' OR Status = 'Y')"
        strCond = strCond & " AND Status IN ('C', 'Y')"
    Else
        strCond = strCond & " AND (Status NOT IN ('C', 'Y'))"
    End If
    
    strCond = strCond & " AND CustomerCode IS NOT NULL"
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
    ElseIf cboCustType.ListIndex = 7 Then
        strCond = strCond & " AND CustTypeFlag in ('M','C','S') "
    ElseIf cboCustType.ListIndex = 8 Then
        strCond = strCond & " AND CustTypeFlag in ('M') "
    End If
    strCond = strCond & strOptional
    
    Set colSearchResult_ = tmpS.CustomerSearch(strCond)
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        If ShowSubContact_ = "Y" Then
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", IIf(tmpItem.SubContact = "", tmpItem.CustomerName, tmpItem.CustomerName & " (" & tmpItem.SubContact & ")"))
        Else
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.CustomerName)
        End If
        tmpList.SubItems(1) = Trim(tmpItem.CustomerCode)
        tmpList.SubItems(2) = Trim(tmpItem.RegisterNo)
        tmpList.SubItems(3) = Trim(tmpItem.UnderSaleName)
        tmpList.SubItems(4) = Trim(tmpItem.ProvinceName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        If firstActive = True Then
            MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation
'            result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลลูกค้าใหม่หรือไม่", vbQuestion + vbYesNo)
'            If result = vbYes Then
'                frmCustomer.cmdSearch.Enabled = False
'                frmCustomer.Show 1
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
'            End If
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
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(F) RDF Waste"
        .AddItem "(L) (p)"
        .AddItem "(T) (k)"
        .AddItem "(G) (e)"
        .AddItem "(R)"
        .AddItem "(M)"
        .AddItem "(N)"
        .ListIndex = 0
    End With
End Sub

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
    cmdOpen_Click
    Unload Me
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

Private Sub txtSale_KeyPress(KeyAscii As Integer)
    If txtSale.Text = "*" Then Set curStaff_ = New BWGCompanyStaff
    If KeyAscii = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If txtSale.Text = "*" Then Set curStaff_ = New BWGCompanyStaff: Exit Sub
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub
