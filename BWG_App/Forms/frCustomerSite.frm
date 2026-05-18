VERSION 5.00
Begin VB.Form frmCustomerSite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "สถานที่ขนกากของเสียของลูกค้า - F027"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDay 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1530
      TabIndex        =   47
      Top             =   3375
      Width           =   1185
   End
   Begin VB.TextBox TxtHour 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1530
      TabIndex        =   46
      Top             =   3780
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rate น้ำมัน"
      Height          =   1905
      Left            =   4410
      TabIndex        =   31
      Top             =   2610
      Width           =   4110
      Begin VB.TextBox TxtTrail 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         TabIndex        =   48
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox Txt10 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2790
         TabIndex        =   42
         Top             =   1035
         Width           =   1185
      End
      Begin VB.TextBox TxtPickup 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         TabIndex        =   40
         Top             =   1035
         Width           =   1185
      End
      Begin VB.TextBox Txt24Q 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2805
         TabIndex        =   38
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox Txt15Q 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         TabIndex        =   36
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox TxtTrailer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2805
         TabIndex        =   34
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox TxtSingle 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   870
         TabIndex        =   32
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "เทรลเลอร์"
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   49
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "สิบล้อ"
         Height          =   195
         Index           =   9
         Left            =   2370
         TabIndex        =   43
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "กระบะ"
         Height          =   195
         Index           =   8
         Left            =   375
         TabIndex        =   41
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "24Q"
         Height          =   195
         Index           =   7
         Left            =   2460
         TabIndex        =   39
         Top             =   630
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "15Q"
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   37
         Top             =   630
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "พ่วง"
         Height          =   195
         Index           =   5
         Left            =   2475
         TabIndex        =   35
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "เดี่ยว"
         Height          =   195
         Index           =   4
         Left            =   450
         TabIndex        =   33
         Top             =   225
         Width           =   375
      End
   End
   Begin VB.TextBox TxtOilNum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      TabIndex        =   28
      Top             =   2565
      Width           =   1455
   End
   Begin VB.TextBox txtLengthNum 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5580
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox chkAddr 
      Caption         =   "ใช้ที่อยู่นี้เป็นที่อยู่ในใบกำกับ (Manifest)"
      Height          =   285
      Left            =   945
      TabIndex        =   10
      Top             =   3000
      Width           =   3075
   End
   Begin VB.TextBox txtContactName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   2130
      Width           =   3195
   End
   Begin VB.TextBox txtFaxNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5580
      TabIndex        =   8
      Top             =   1740
      Width           =   2895
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   1740
      Width           =   3195
   End
   Begin VB.TextBox txtZipCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5580
      TabIndex        =   6
      Top             =   1350
      Width           =   1455
   End
   Begin VB.TextBox txtSiteName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   570
      Width           =   3195
   End
   Begin VB.TextBox txtAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5550
      TabIndex        =   2
      Top             =   570
      Width           =   2925
   End
   Begin VB.TextBox txtTumbol 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   2835
   End
   Begin VB.TextBox txtAumphur 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5580
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtProvince 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2835
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Enabled         =   0   'False
      Height          =   795
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4635
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4635
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4635
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4635
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ชั่วโมงที่วิ่ง :"
      Height          =   195
      Left            =   675
      TabIndex        =   45
      Top             =   3825
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "จำนวนวันที่วิ่ง :"
      Height          =   195
      Left            =   405
      TabIndex        =   44
      Top             =   3420
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "น้ำมัน"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   30
      Top             =   2610
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ลิตร"
      Height          =   255
      Index           =   2
      Left            =   2460
      TabIndex        =   29
      Top             =   2610
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Km"
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   27
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ระยะทาง"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   26
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้ติดต่อ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   24
      Top             =   2190
      Width           =   705
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   4530
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แฟกซ์"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   5070
      TabIndex        =   23
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "โทร"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   600
      TabIndex        =   22
      Top             =   1800
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสไปรษณีย์"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   4560
      TabIndex        =   21
      Top             =   1410
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อสถานที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   20
      Top             =   630
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่, หมู่, ถนน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4380
      TabIndex        =   19
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำบล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5070
      TabIndex        =   17
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "จังหวัด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   16
      Top             =   1410
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   300
      TabIndex        =   15
      Top             =   210
      Width           =   570
   End
End
Attribute VB_Name = "frmCustomerSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myCustSite_ As BWGCustomerSite
Dim curTumbol_ As IMITumbol
Dim colCompanyGroup_ As Collection
Dim curCustomer_ As BWGCustomer
Dim callerForm_ As Form

Private Sub chkAddr_Validate(Cancel As Boolean)
    If chkAddr.Value = 1 Then
        myCustSite_.isManifestAddr = True
    Else
        myCustSite_.isManifestAddr = False
    End If
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลสถานที่ขนส่งกากของเสียของลูกค้ารายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myCustSite_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewSite
            txtSiteName.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myCustSite_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustSite_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewSite
    txtSiteName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim tmpStr
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, UpdateRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If

    If Trim(txtSiteName.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อสถานที่", vbExclamation
        txtSiteName.SetFocus
        Exit Sub
    End If

    If Trim(txtTumbol.Text) = "" Then
        MsgBox "กรุณาระบุตำบลของสถานที่", vbExclamation
        txtTumbol.SetFocus
        Exit Sub
    End If

    If myCustSite_.isChanged Then
        myCustSite_.Save
    End If
    If myCustSite_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
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
    Set curTumbol_ = New IMITumbol
    CreateNewSite
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtSiteName.Text = ""
    txtAddr.Text = ""
    txtTumbol.Text = ""
    txtAumphur.Text = ""
    txtProvince.Text = ""
    txtZipCode.Text = ""
    txtTelNo.Text = ""
    txtFaxNo.Text = ""
    txtContactName.Text = ""
    chkAddr.Value = 0
    txtLengthNum.Text = ""
    TxtOilNum.Text = ""
    Txt10.Text = ""
    Txt15Q.Text = ""
    Txt24Q.Text = ""
    TxtTrailer.Text = ""
    TxtTrail.Text = ""
    TxtPickup.Text = ""
    TxtSingle.Text = ""
    TxtDay.Text = ""
    TxtHour.Text = ""
End Sub
Private Sub CreateNewSite()
    Set myCustSite_ = Nothing
    Set myCustSite_ = New BWGCustomerSite
    With myCustSite_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curCustomer_ Is Nothing Then .CustomerID = curCustomer_.ID
        .isChanged = False
    End With
End Sub

Public Sub SetDataCustomerSite(selSite As BWGCustomerSite)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myCustSite_ = Nothing
    Set myCustSite_ = selSite
    If myCustSite_ Is Nothing Then Exit Sub
    With myCustSite_
        txtSiteName.Text = .SiteName
        txtAddr.Text = .Address
        Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
        If tmpCol.Count > 0 Then SetDataTumbol tmpCol(1), False
        txtZipCode.Text = .ZipCode
        txtTelNo.Text = .TelNo
        txtFaxNo.Text = .FaxNo
        txtContactName.Text = .ContactName
        txtLengthNum.Text = .LengthNum
        If CDbl(.RateSingle) > 0 Then
            TxtSingle.Text = .RateSingle
        End If
        If CDbl(.RateTrailer) > 0 Then
            TxtTrailer.Text = .RateTrailer
        End If
        If CDbl(.RateTrail) > 0 Then
            TxtTrail.Text = .RateTrail
        End If
        If CDbl(.Rate15Q) > 0 Then
            Txt15Q.Text = .Rate15Q
        End If
        If CDbl(.Rate24Q) > 0 Then
            Txt24Q.Text = .Rate24Q
        End If
        If CDbl(.RatePickup) > 0 Then
            TxtPickup.Text = .RatePickup
        End If
        If CDbl(.Rate10) > 0 Then
            Txt10.Text = .Rate10
        End If
        TxtOilNum.Text = .OilNum
        If .RunDay = "0" Then
            TxtDay.Text = ""
        Else
            TxtDay.Text = .RunDay
        End If
        If .RunHour = "0" Then
            TxtHour.Text = ""
        Else
            TxtHour.Text = .RunHour
        End If
        If .isManifestAddr Then
            chkAddr.Value = 1
        Else
            chkAddr.Value = 0
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myCustSite_ Is Nothing Then Exit Sub
    If myCustSite_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustSite_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
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
            myCustSite_.AumphurID = .AumphurID
            myCustSite_.AumphurName = .AumphurName
            myCustSite_.TumbolID = .ID
            myCustSite_.TumbolName = .Name
            myCustSite_.ProvinceID = .ProvinceID
            myCustSite_.ProvinceName = .ProvinceName
        End If
    End With
End Sub
Private Sub Txt10_Validate(Cancel As Boolean)
    myCustSite_.Rate10 = Trim(Txt10.Text)
End Sub

Private Sub Txt15Q_Validate(Cancel As Boolean)
    myCustSite_.Rate15Q = Trim(Txt15Q.Text)
End Sub

Private Sub Txt24Q_Validate(Cancel As Boolean)
    myCustSite_.Rate24Q = Trim(Txt24Q.Text)
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
    myCustSite_.Address = Trim(txtAddr.Text)
End Sub

Private Sub txtContactName_Validate(Cancel As Boolean)
    myCustSite_.ContactName = Trim(txtContactName.Text)
End Sub
Private Sub TxtDay_Validate(Cancel As Boolean)
    myCustSite_.RunDay = Trim(TxtDay.Text)
End Sub

Private Sub txtFaxNo_Validate(Cancel As Boolean)
    myCustSite_.FaxNo = Trim(txtFaxNo.Text)
End Sub
Private Sub TxtHour_Validate(Cancel As Boolean)
        myCustSite_.RunHour = Trim(TxtHour.Text)
End Sub
Private Sub txtLengthNum_Validate(Cancel As Boolean)
        myCustSite_.LengthNum = Trim(txtLengthNum.Text)
End Sub
Private Sub TxtOilNum_Validate(Cancel As Boolean)
    myCustSite_.OilNum = Trim(TxtOilNum.Text)
End Sub
Private Sub TxtPickup_Validate(Cancel As Boolean)
    myCustSite_.RatePickup = Trim(TxtPickup.Text)
End Sub
Private Sub TxtSingle_Validate(Cancel As Boolean)
    myCustSite_.RateSingle = Trim(TxtSingle.Text)
End Sub

Private Sub txtSiteName_Validate(Cancel As Boolean)
    myCustSite_.SiteName = Trim(txtSiteName.Text)
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    myCustSite_.TelNo = Trim(txtTelNo.Text)
End Sub

Private Sub TxtTrail_Validate(Cancel As Boolean)
    myCustSite_.RateTrail = Trim(TxtTrail.Text)
End Sub

Private Sub TxtTrailer_Validate(Cancel As Boolean)
    myCustSite_.RateTrailer = Trim(TxtTrailer.Text)
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With myCustSite_
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
            With myCustSite_
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
            If myCustSite_.TumbolID = "" Then txtTumbol.Text = ""
        End If
    End If
End Sub

Private Sub txtZipCode_Validate(Cancel As Boolean)
    myCustSite_.ZipCode = Trim(txtZipCode.Text)
End Sub

Public Sub SetCallerForm(selFrm As Form, selCust As BWGCustomer)
    Set callerForm_ = selFrm
    Set curCustomer_ = selCust
    txtCustomer.Text = selCust.CustomerName
    If Not myCustSite_ Is Nothing Then myCustSite_.CustomerID = curCustomer_.ID
End Sub
