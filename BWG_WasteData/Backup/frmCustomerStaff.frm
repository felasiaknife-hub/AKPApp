VERSION 5.00
Begin VB.Form frmCustomerStaff 
   Caption         =   "รายชื่อเจ้าหน้าที่ของลูกค้า - F028"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   22
      Top             =   1800
      Width           =   6885
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   19
      Top             =   1395
      Width           =   2925
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5130
      TabIndex        =   18
      Top             =   1395
      Width           =   2925
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2220
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2220
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2220
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2220
      Width           =   1005
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   180
      Width           =   6885
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5130
      TabIndex        =   5
      Top             =   570
      Width           =   2925
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   4
      Top             =   570
      Width           =   2925
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5130
      TabIndex        =   3
      Top             =   990
      Width           =   2925
   End
   Begin VB.TextBox txtFaxNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   2475
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   2265
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox txtDepartment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Top             =   990
      Width           =   2925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ที่อยู่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   720
      TabIndex        =   23
      Top             =   1875
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   21
      Top             =   1470
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เรื่องที่ติดต่อ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4245
      TabIndex        =   20
      Top             =   1455
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   510
      TabIndex        =   17
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "นามสกุล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4470
      TabIndex        =   16
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อพนักงาน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   15
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "โทร"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   4815
      TabIndex        =   14
      Top             =   1050
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แฟกซ์"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   13
      Top             =   2535
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   0
      Top             =   2160
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำแหน่ง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   510
      TabIndex        =   12
      Top             =   2325
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แผนก"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   705
      TabIndex        =   11
      Top             =   1065
      Width           =   450
   End
End
Attribute VB_Name = "frmCustomerStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myCustStaff_ As BWGCustomerStaff
Dim curCustomer_ As BWGCustomer
Dim callerForm_ As Form

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลชื่อเจ้าหน้าที่ของลูกค้ารายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myCustStaff_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewCustStaff
            txtFName.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
    frmCustomer.ShowCustStaff
    Unload Me
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myCustStaff_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCustStaff_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewCustStaff
    txtFName.SetFocus
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

    If Trim(txtFName.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อเจ้าหน้าที่ก่อนจัดเก็บข้อมูล", vbExclamation
        txtFName.SetFocus
        Exit Sub
    End If

    If myCustStaff_.isChanged Then
        myCustStaff_.Save
        frmCustomer.ShowCustStaff
    End If
    Unload Me
    If myCustStaff_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtFName.Text = ""
    txtLName.Text = ""
    txtPosition.Text = ""
    txtDepartment.Text = ""
    txtTelNo.Text = ""
    txtFaxNo.Text = ""
    txtMail.Text = ""
    txtSubject.Text = ""
End Sub

Private Sub CreateNewCustStaff()
    Set myCustStaff_ = Nothing
    Set myCustStaff_ = New BWGCustomerStaff
    With myCustStaff_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curCustomer_ Is Nothing Then .CustomerID = curCustomer_.ID
        .isChanged = False
    End With
End Sub

Public Sub SetDataCustomerStaff(selSite As BWGCustomerStaff)
    Set myCustStaff_ = Nothing
    Set myCustStaff_ = selSite
    If myCustStaff_ Is Nothing Then Exit Sub
    With myCustStaff_
        txtFName.Text = .FirstName
        txtLName.Text = .LastName
        txtPosition.Text = .PositionName
        txtDepartment.Text = .DeptName
        txtTelNo.Text = .TelNo
        txtFaxNo.Text = .FaxNo
        txtMail.Text = .Email
        txtSubject.Text = .Subjact
        TxtAddress.Text = .Address
    End With
End Sub

Private Sub Form_Load()
'    CanAccess = False: CanUpdate = False
'    SetPermission Me, CanAccess, CanUpdate
'    If CanAccess = False Then
'        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
'        Unload Me
'    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    CreateNewCustStaff

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myCustStaff_ Is Nothing Then Exit Sub
    If myCustStaff_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCustStaff_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetCallerForm(selFrm As Form, selCust As BWGCustomer)
    Set callerForm_ = selFrm
    Set curCustomer_ = selCust
    txtCustomer.Text = selCust.CustomerName
    If Not myCustStaff_ Is Nothing Then myCustStaff_.CustomerID = curCustomer_.ID
End Sub

Private Sub TxtAddress_Validate(Cancel As Boolean)
    myCustStaff_.Address = Trim(TxtAddress.Text)
End Sub

Private Sub txtDepartment_Validate(Cancel As Boolean)
    myCustStaff_.DeptName = Trim(txtDepartment.Text)
End Sub

Private Sub txtFaxNo_Validate(Cancel As Boolean)
    myCustStaff_.FaxNo = Trim(txtFaxNo.Text)
End Sub

Private Sub txtFName_Validate(Cancel As Boolean)
    myCustStaff_.FirstName = Trim(txtFName.Text)
End Sub

Private Sub txtLName_Validate(Cancel As Boolean)
    myCustStaff_.LastName = Trim(txtLName.Text)
End Sub
Private Sub txtMail_Validate(Cancel As Boolean)
    myCustStaff_.Email = Trim(txtMail.Text)
End Sub

Private Sub txtPosition_Validate(Cancel As Boolean)
    myCustStaff_.PositionName = Trim(txtPosition.Text)
End Sub

Private Sub txtSubject_Validate(Cancel As Boolean)
        myCustStaff_.Subjact = Trim(txtSubject.Text)
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    myCustStaff_.TelNo = Trim(txtTelNo.Text)
End Sub

