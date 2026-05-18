VERSION 5.00
Begin VB.Form frmContract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "สัญญาการให้บริการการขนกากของเสีย - [F021]"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WasteDataForm.ctlDate dtEnd 
      Height          =   315
      Left            =   4260
      TabIndex        =   15
      Top             =   960
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteDataForm.ctlDate dtStart 
      Height          =   315
      Left            =   1500
      TabIndex        =   14
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   570
      Width           =   4515
   End
   Begin VB.TextBox txtContractNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7080
      TabIndex        =   2
      Top             =   570
      Width           =   1515
   End
   Begin VB.TextBox txtRemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   1350
      Width           =   7095
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2070
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2070
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2070
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2070
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ทำสัญญากับบริษัท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   180
      TabIndex        =   13
      Top             =   630
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่สัญญา"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6210
      TabIndex        =   12
      Top             =   630
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่สิ้นสุด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3450
      TabIndex        =   11
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เริ่ม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   1020
      Width           =   585
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   1920
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   780
      TabIndex        =   9
      Top             =   1410
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   840
      TabIndex        =   8
      Top             =   210
      Width           =   570
   End
End
Attribute VB_Name = "frmContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim myContract_ As BWGContract
Dim curCustomer_ As BWGCustomer
Dim curCompany_ As BWGCompanyConfig
Dim callerForm_ As Form

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลสัญญาของลูกค้ารายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myContract_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewContract
            txtCompany.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myContract_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myContract_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewContract
    txtCompany.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If

    If Trim(txtCompany.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        txtCompany.SetFocus
        Exit Sub
    End If

    If Trim(txtContractNo.Text) = "" Then
        MsgBox "กรุณาระบุเลขที่สัญญา", vbExclamation
        txtContractNo.SetFocus
        Exit Sub
    End If
    If dtStart.ValueYMD = dtEnd.ValueYMD Then
        MsgBox "ข้อมูลวันเริ่มต้นและสิ้นสุดสัญญาไม่ถูกต้อง จะต้องไม่เป็นวันเดียวกัน", vbExclamation
        dtEnd.SetFocus
        Exit Sub
    End If

    If myContract_.isChanged Then
        myContract_.Save
    End If
    If myContract_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub dtEnd_Validate(Cancel As Boolean)
    myContract_.EndDate = Trim(dtEnd.ValueYMD)
End Sub

Private Sub dtStart_Validate(Cancel As Boolean)
    myContract_.StartDate = Trim(dtStart.ValueYMD)
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    CreateNewContract
    Set curCompany_ = New BWGCompanyConfig
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtContractNo.Text = ""
    dtStart.ValueYMD = TodayDate
    dtEnd.ValueYMD = TodayDate
    txtRemDesc.Text = ""
End Sub

Private Sub CreateNewContract()
    Set myContract_ = Nothing
    Set myContract_ = New BWGContract
    With myContract_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        If Not curCompany_ Is Nothing Then
            .companyID = curCompany_.ID
            .CompanyName = curCompany_.CompanyNameTH
        End If
        .StartDate = dtStart.ValueYMD
        .EndDate = dtEnd.ValueYMD
        .isChanged = False
    End With
End Sub

Public Sub SetDataContract(selSite As BWGContract)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myContract_ = Nothing
    Set myContract_ = selSite
    If myContract_ Is Nothing Then Exit Sub
    With myContract_
        Set tmpCol = tmpS.CompanySearch("CompanyID='" & .companyID & "'")
        If tmpCol.Count > 0 Then SetDataCompany tmpCol(1), False
        txtContractNo.Text = .docNo
        dtStart.ValueYMD = .StartDate
        dtEnd.ValueYMD = .EndDate
        txtRemDesc.Text = .RemDesc
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myContract_ Is Nothing Then Exit Sub
    If myContract_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myContract_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetDataCompany(selCompany As BWGCompanyConfig, Optional isAssigned As Boolean = True)
    Set curCompany_ = Nothing
    Set curCompany_ = selCompany
    If curCompany_ Is Nothing Then Exit Sub
    With curCompany_
        txtCompany.Text = .CompanyNameTH
        If isAssigned Then
            myContract_.companyID = .ID
            myContract_.CompanyName = .CompanyNameTH
        End If
    End With
End Sub

Public Sub SetCallerForm(selFrm As Form, selCust As BWGCustomer)
    Set callerForm_ = selFrm
    Set curCustomer_ = selCust
    txtCustomer.Text = selCust.CustomerName
    If Not myContract_ Is Nothing Then myContract_.CustomerID = curCustomer_.ID
End Sub

Private Sub txtCompany_Validate(Cancel As Boolean)
    If Trim(txtCompany.Text) = "" And Trim(curCompany_.CompanyNameTH) <> Trim(txtCompany.Text) Then
        Set curCompany_ = Nothing
        Set curCompany_ = New BWGCompanyConfig
        With myContract_
            .companyID = ""
            .CompanyName = ""
        End With
    Else
        If Trim(curCompany_.CompanyNameTH) <> Trim(txtCompany.Text) Then
            Set curCompany_ = Nothing
            Set curCompany_ = New BWGCompanyConfig
            With myContract_
                .companyID = ""
                .CompanyName = ""
            End With
'            Load frCompanySearch
'            With frCompanySearch
'                .getCriteria (txtCompany.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
            If myContract_.companyID = "" Then txtCompany.Text = ""
        End If
    End If
End Sub

Private Sub txtContractNo_Validate(Cancel As Boolean)
    myContract_.docNo = Trim(txtContractNo.Text)
End Sub

Private Sub txtRemDesc_Validate(Cancel As Boolean)
    myContract_.RemDesc = Trim(txtRemDesc.Text)
End Sub
