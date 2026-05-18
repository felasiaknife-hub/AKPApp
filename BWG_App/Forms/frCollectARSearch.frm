VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoiceSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F011"
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
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1620
      Width           =   4275
   End
   Begin WasteManagment.ctlDate dtFrom 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Text            =   "*"
      Top             =   1260
      Width           =   4275
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5910
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5910
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   30
      Picture         =   "frCollectARSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5910
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3465
      Left            =   0
      TabIndex        =   4
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
         Text            =   "เลขที่ใบวางบิล"
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
   Begin WasteManagment.ctlDate dtTo 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   900
      Width           =   1815
      _ExtentX        =   3201
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
      TabIndex        =   13
      Top             =   960
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
      TabIndex        =   12
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใส่เลขวางบิลเท่านั้น"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   570
      TabIndex        =   11
      Top             =   1320
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
      Caption         =   "ค้นหาใบวางบิล"
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
      TabIndex        =   10
      Top             =   75
      Width           =   1320
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
      TabIndex        =   9
      Top             =   2130
      Width           =   9285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า, รหัสลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frCollectARSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmInvoiceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean
Dim curCustomer_ As BWGCustomer

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
    callerForm_.SetDataBillHeader colSearchResult_(lvResult.SelectedItem.Key)
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
            callerForm_.SetDataBillHeader colSearchResult_(1)
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
Dim x%, result%, Y%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGBillingHeader
    If Trim(txtCriteria.Text) = "" Then txtCriteria.Text = "*"
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
    
    If Not curCustomer_ Is Nothing Then
        strCond = "(CustomerID = '" & curCustomer_.ID & "') AND "
    End If
    If criStr <> "%" Then
        strCond = strCond & "(BillingNo Like '%" & criStr & "%') AND "
    End If
    If firstActive = False Then
        strCond = strCond & "  (IssuedDate>='" & dtFrom.ValueYMD & "' And IssuedDate<='" & dtTo.ValueYMD & "')"
    End If
    strCond = SelectWithCompanyID(strCond, True)
    Set colSearchResult_ = tmpS.BillingHeaderSearch(strCond)
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
            If tmpItem.BillingNo <> "" Then
                Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.BillingNo)
            Else
                Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.PreInvNo)
            End If
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
        ElseIf tmpItem.isCustAccepted Then
            tmpList.SubItems(3) = "ลูกค้ารับวางบิลแล้ว"
            tmpList.ForeColor = &H8000&
            tmpList.Bold = True
            For Y = 1 To 3
                tmpList.ListSubItems(Y).Bold = True
                tmpList.ListSubItems(Y).ForeColor = &H8000&
            Next
        Else
            tmpList.SubItems(3) = "New"
        End If
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpS = Nothing
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOpen_Click
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub txtCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate True
End Sub

Private Sub txtCustomer_LostFocus()
    If Trim(txtCustomer.Text) = "*" Or txtCustomer.Text = "" Then
        Set curCustomer_ = Nothing
    End If
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
Dim selStr As String
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    selStr = Trim(curCustomer_.CustomerName)
    If Trim(txtCustomer.Text) = "" And Trim(selStr) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(selStr) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .ShowSubContact = True
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal ""
'                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_ Is Nothing Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = selCust
    If Not curCustomer_ Is Nothing Then
        txtCustomer.Text = curCustomer_.CustomerName
    Else
        txtCustomer.Text = ""
    End If
End Sub
