VERSION 5.00
Begin VB.Form frmAccountPayable 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payable"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtVenderCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7320
      TabIndex        =   34
      Text            =   "== AUTO =="
      Top             =   420
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   60
      TabIndex        =   11
      Top             =   2550
      Width           =   9270
      Begin VB.TextBox txtCompanyRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1410
         TabIndex        =   10
         Top             =   1380
         Width           =   4935
      End
      Begin VB.TextBox txtWebsite 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1410
         TabIndex        =   9
         Top             =   1005
         Width           =   4935
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         Top             =   615
         Width           =   4935
      End
      Begin VB.TextBox txtTelNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1410
         TabIndex        =   6
         Top             =   225
         Width           =   2295
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4350
         TabIndex        =   7
         Top             =   225
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   675
         TabIndex        =   24
         Top             =   1395
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เว็บไซต์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   735
         TabIndex        =   22
         Top             =   1065
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อีเมลล์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   855
         TabIndex        =   21
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แฟกซ์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   3825
         TabIndex        =   20
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรศัพท์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   19
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "ข้อมูล Vender"
      Height          =   1650
      Left            =   30
      TabIndex        =   17
      Top             =   855
      Width           =   9315
      Begin VB.TextBox txtLicenseNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7470
         TabIndex        =   5
         Top             =   1050
         Width           =   1695
      End
      Begin VB.TextBox txtTumbol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3900
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.TextBox txtProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6450
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.TextBox txtZipCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   1050
         Width           =   1185
      End
      Begin VB.TextBox txtTaxID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4410
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1050
         Width           =   1335
      End
      Begin VB.TextBox txtCompanyNameTH 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   0
         Top             =   270
         Width           =   3165
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5910
         TabIndex        =   1
         Top             =   270
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขทะเบียนบริษัท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5940
         TabIndex        =   33
         Top             =   1110
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อบริษัท/โรงงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   32
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่, หมู่, ถนน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4770
         TabIndex        =   31
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตำบล"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อำเภอ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   29
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   5940
         TabIndex        =   28
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสไปรษณีย์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   420
         TabIndex        =   27
         Top             =   1110
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัวผู้เสียภาษี"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   2745
         TabIndex        =   23
         Top             =   1140
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   7215
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4605
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4605
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4605
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4605
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4605
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   6480
      TabIndex        =   35
      Top             =   480
      Width           =   750
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9375
      Y1              =   4515
      Y2              =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูลเจ้าหนี้การค้าทั้วไป"
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
      Index           =   1
      Left            =   180
      TabIndex        =   18
      Top             =   0
      Width           =   2115
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   5800
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5680
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frmAccountPayable.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9345
   End
End
Attribute VB_Name = "frmAccountPayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myPayable_ As BWGTransportVendor
Dim colSuppilerContact_ As Collection
Dim curTumbol_ As IMITumbol

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลกลุ่มย่อยของสินค้านี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myPayable_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewTransportVender
            txtCompanyNameTH.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myPayable_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myPayable_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewTransportVender
    txtCompanyNameTH.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload frmTransportVenderSearch
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อบริษัทภาษาไทย", vbExclamation
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If

    If myPayable_.isChanged Then
        myPayable_.Save
        txtVenderCode.Text = myPayable_.VendorCode
    End If
    If myPayable_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    If myPayable_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myPayable_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmTransportVenderSearch
    With frmTransportVenderSearch
        .SetCallerForm Me
        .VendorType "GL"
        .Show 1
    End With
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
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    Set myPayable_ = Nothing
    Set myPayable_ = New BWGTransportVendor

    CreateNewTransportVender
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grDisposeNo" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtCompanyNameTH.Text = ""
    txtAddr.Text = ""
    txtTumbol.Text = ""
    txtAumphur.Text = ""
    txtProvince.Text = ""
    txtZipCode.Text = ""
    txtTaxID.Text = ""
    txtTelNo.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtWebsite.Text = ""
    txtCompanyRem.Text = ""
    txtLicenseNo.Text = ""
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    Set myPayable_ = Nothing
    Set myPayable_ = New BWGTransportVendor

End Sub

Private Sub CreateNewTransportVender()
Set myPayable_ = Nothing
Set myPayable_ = New BWGTransportVendor
    With myPayable_
        .ID = GetGUID
        .VendorType = "GL"
        .CreatedDate = TodayDate
        .isDeleted = False
        .isChanged = False
    End With
    'SetEnabledScreen
End Sub

Public Sub SetDataTransportVender(selCompany As BWGTransportVendor)
Dim x%, tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myPayable_ = Nothing
    Set myPayable_ = selCompany
    If myPayable_ Is Nothing Then Exit Sub
    With myPayable_
        txtVenderCode.Text = myPayable_.VendorCode
        txtCompanyNameTH.Text = .NameTH
        txtAddr.Text = .Addr
        txtZipCode.Text = .ZipCode
        txtTaxID.Text = .TaxID
        txtTelNo.Text = .TelNo
        txtFax.Text = .FaxNo
        txtEmail.Text = .EmailAddr
        txtWebsite.Text = .Website
        txtCompanyRem.Text = .SupplierRem
        txtLicenseNo.Text = .TransLicenseNo
        Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
        If tmpCol.Count > 0 Then
            SetDataTumbol tmpCol(1), False
            txtTumbol.Text = .TumbolName
            txtAumphur.Text = .AumphurName
            txtProvince.Text = .ProvinceName
        End If
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myPayable_ Is Nothing Then Exit Sub
    If myPayable_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myPayable_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTransportVenderSearch
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If
    myPayable_.Addr = Trim(txtAddr.Text)
End Sub

Private Sub txtCompanyNameTH_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If
    myPayable_.NameTH = Trim(txtCompanyNameTH.Text)
End Sub

Private Sub txtCompanyRem_Change()
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If
    myPayable_.SupplierRem = Trim(txtCompanyRem.Text)
End Sub


Private Sub txtEmail_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If

    myPayable_.EmailAddr = Trim(txtEmail.Text)
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If

    myPayable_.FaxNo = Trim(txtFax.Text)
End Sub

Private Sub txtLicenseNo_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If

    myPayable_.TransLicenseNo = Trim(txtLicenseNo.Text)
End Sub

Private Sub txtTaxID_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If

    myPayable_.TaxID = Trim(txtTaxID.Text)
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If

    myPayable_.TelNo = Trim(txtTelNo.Text)
End Sub

Private Sub txtWebSite_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If

    myPayable_.Website = Trim(txtWebsite.Text)
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With myPayable_
            .TumbolID = ""
            .TumbolName = ""
            .AumphurID = ""
            .AumphurName = ""
            .ProvinceID = ""
            .ProvinceName = ""
        End With
    Else
        If Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
            Set curTumbol_ = Nothing
            Set curTumbol_ = New IMITumbol
            txtAumphur.Text = ""
            txtProvince.Text = ""
            With myPayable_
                .TumbolID = ""
                .TumbolName = ""
                .AumphurID = ""
                .AumphurName = ""
                .ProvinceID = ""
                .ProvinceName = ""
            End With
            Load frmTumbolSearch
            With frmTumbolSearch
                .getCriteria (txtTumbol.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myPayable_.TumbolID = "" Then txtTumbol.Text = ""
        End If
    End If
End Sub

Public Sub SetDataTumbol(selTumbol As IMITumbol, Optional isAssigned As Boolean = True)
    Set curTumbol_ = Nothing
    Set curTumbol_ = selTumbol
    If curTumbol_ Is Nothing Then Exit Sub
    With curTumbol_
        txtTumbol.Text = .Name
        txtAumphur.Text = .AumphurName
        txtProvince.Text = .ProvinceName
        If isAssigned Then
            myPayable_.AumphurID = .AumphurID
            myPayable_.AumphurName = .AumphurName
            myPayable_.TumbolID = .ID
            myPayable_.TumbolName = .Name
            myPayable_.ProvinceID = .ProvinceID
            myPayable_.ProvinceName = .ProvinceName
        End If
    End With
End Sub

Private Sub txtZipCode_Validate(Cancel As Boolean)
    If myPayable_ Is Nothing Then
        CreateNewTransportVender
    End If
    myPayable_.ZipCode = Trim(txtZipCode.Text)

End Sub
