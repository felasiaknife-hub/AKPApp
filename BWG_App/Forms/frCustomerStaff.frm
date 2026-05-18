VERSION 5.00
Begin VB.Form frmCustomerStaff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "รายชื่อเจ้าหน้าที่ของลูกค้า - F028"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1125
      TabIndex        =   22
      Top             =   1800
      Width           =   6885
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5085
      TabIndex        =   19
      Top             =   1380
      Width           =   2925
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1125
      TabIndex        =   18
      Top             =   1380
      Width           =   2925
   End
   Begin VB.TextBox txtDepartment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   4
      Top             =   960
      Width           =   2925
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   2220
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox txtFaxNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   2610
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5100
      TabIndex        =   5
      Top             =   960
      Width           =   2925
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   570
      Width           =   2925
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5100
      TabIndex        =   2
      Top             =   570
      Width           =   2925
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   6885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2265
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2265
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2265
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2265
      Width           =   1005
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
      Caption         =   "เรื่องที่ติดต่อ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4200
      TabIndex        =   21
      Top             =   1450
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMail"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   675
      TabIndex        =   20
      Top             =   1450
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แผนก"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   630
      TabIndex        =   17
      Top             =   1020
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำแหน่ง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   2205
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แฟกซ์"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   630
      TabIndex        =   15
      Top             =   2670
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "โทร"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   4740
      TabIndex        =   14
      Top             =   1005
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อพนักงาน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   13
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "นามสกุล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   570
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
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myCustStaff_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
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

Private Sub CmdSave_Click()
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

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    CreateNewCustStaff
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
    TxtPosition.Text = ""
    txtDepartment.Text = ""
    txtTelNo.Text = ""
    txtFaxNo.Text = ""
    TxtMail.Text = ""
    txtSubject.Text = ""
    TxtAddress.Text = ""
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
        TxtPosition.Text = .PositionName
        txtDepartment.Text = .DeptName
        txtTelNo.Text = .TelNo
        txtFaxNo.Text = .FaxNo
        TxtMail.Text = .Email
        txtSubject.Text = .Subjact
        TxtAddress.Text = .Address
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myCustStaff_ Is Nothing Then Exit Sub
    If myCustStaff_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
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

Private Sub TxtAddress_Change()
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
    myCustStaff_.Email = Trim(TxtMail.Text)
End Sub

Private Sub TxtPosition_Validate(Cancel As Boolean)
    myCustStaff_.PositionName = Trim(TxtPosition.Text)
End Sub

Private Sub txtSubject_Validate(Cancel As Boolean)
    myCustStaff_.Subjact = Trim(txtSubject.Text)
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    myCustStaff_.TelNo = Trim(txtTelNo.Text)
End Sub
