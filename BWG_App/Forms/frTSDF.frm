VERSION 5.00
Begin VB.Form frmTSDF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ข้อมูลบริษัท TSDF - F135"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLandFill 
      Caption         =   "ทำฝังกลบได้"
      Height          =   255
      Left            =   6330
      TabIndex        =   11
      Top             =   1770
      Width           =   1275
   End
   Begin VB.CheckBox chkBlending 
      Caption         =   "ทำ Blending ได้"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   1770
      Width           =   1485
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2430
      Width           =   1005
   End
   Begin VB.TextBox txtContactName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   1740
      Width           =   3345
   End
   Begin VB.TextBox txtFaxNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5730
      TabIndex        =   8
      Top             =   1350
      Width           =   2895
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   1350
      Width           =   3345
   End
   Begin VB.TextBox txtZipCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7170
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtTSDFNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6780
      TabIndex        =   1
      Top             =   180
      Width           =   1845
   End
   Begin VB.TextBox txtAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   570
      Width           =   3975
   End
   Begin VB.TextBox txtTumbol 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5730
      TabIndex        =   3
      Top             =   570
      Width           =   2895
   End
   Begin VB.TextBox txtAumphur 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtProvince 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3750
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtCompanyName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   4725
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2430
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2430
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2430
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2430
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อผู้ติดต่อ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   26
      Top             =   1800
      Width           =   705
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   2310
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แฟกซ์"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   5220
      TabIndex        =   25
      Top             =   1410
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
      TabIndex        =   24
      Top             =   1410
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสไปรษณีย์"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   6150
      TabIndex        =   23
      Top             =   1020
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขประจำตัว"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   5820
      TabIndex        =   22
      Top             =   240
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ที่อยู่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   21
      Top             =   630
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำบล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   5250
      TabIndex        =   20
      Top             =   630
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   450
      TabIndex        =   19
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
      Left            =   3210
      TabIndex        =   18
      Top             =   1020
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อบริษัท"
      Height          =   195
      Index           =   24
      Left            =   270
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmTSDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myTSDF_ As BWGTSDF
Dim curTumbol_ As IMITumbol

Private Sub chkBlending_Validate(Cancel As Boolean)
    myTSDF_.CanBlending = chkBlending.Value = 1
End Sub

Private Sub chkLandFill_Validate(Cancel As Boolean)
    myTSDF_.CanLandFill = chkLandFill.Value = 1
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูล TSDF รายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myTSDF_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewTSDF
            txtCompanyName.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myTSDF_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myTSDF_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewTSDF
    txtCompanyName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtCompanyName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท TSDF", vbExclamation
        txtCompanyName.SetFocus
        Exit Sub
    End If

    If Trim(txtTSDFNo.Text) = "" Then
        MsgBox "กรุณาใส่เลขประจำตัว TSDF", vbExclamation
        txtTSDFNo.SetFocus
        Exit Sub
    End If

    If Trim(txtTumbol.Text) = "" Then
        MsgBox "กรุณาระบุตำบลของสถานที่", vbExclamation
        txtTumbol.SetFocus
        Exit Sub
    End If

    If myTSDF_.isChanged Then
        myTSDF_.Save
    End If
    If myTSDF_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim myRes%
    If myTSDF_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            cmdSave_Click
            If myTSDF_.isChanged Then
                Exit Sub
            End If
        ElseIf myRes = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmTSDFSearch
    With frmTSDFSearch
        .SetCallerForm Me
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
        cmdClear.Picture = .picNew.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Set curTumbol_ = New IMITumbol
    CreateNewTSDF
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtCompanyName.Text = ""
    txtTSDFNo.Text = ""
    txtAddr.Text = ""
    txtTumbol.Text = ""
    txtAumphur.Text = ""
    txtProvince.Text = ""
    txtZipCode.Text = ""
    txtTelNo.Text = ""
    txtFaxNo.Text = ""
    txtContactName.Text = ""
    chkBlending.Value = 0
    chkLandFill.Value = 0
End Sub

Private Sub CreateNewTSDF()
    Set myTSDF_ = Nothing
    Set myTSDF_ = New BWGTSDF
    With myTSDF_
        .ID = GetGUID ' getNewProductID("CC")
        .CanBlending = False
        .CanLandFill = False
        .isChanged = False
    End With
End Sub

Public Sub SetDataTSDF(selSite As BWGTSDF)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myTSDF_ = Nothing
    Set myTSDF_ = selSite
    If myTSDF_ Is Nothing Then Exit Sub
    With myTSDF_
        txtCompanyName.Text = .CompanyName
        txtTSDFNo.Text = .TSDFRegisNo
        txtAddr.Text = .Addr1
        Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
        If tmpCol.Count > 0 Then SetDataTumbol tmpCol(1), False
        txtZipCode.Text = .ZipCode
        txtTelNo.Text = .TelNo
        txtFaxNo.Text = .FaxNo
        txtContactName.Text = .ContactName
        If .CanBlending Then
            chkBlending.Value = 1
        Else
            chkBlending.Value = 0
        End If
        If .CanLandFill Then
            chkLandFill.Value = 1
        Else
            chkLandFill.Value = 0
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myTSDF_ Is Nothing Then Exit Sub
    If myTSDF_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myTSDF_.isChanged Then Cancel = 1
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
            myTSDF_.AumphurID = .AumphurID
            myTSDF_.AumphurName = .AumphurName
            myTSDF_.TumbolID = .ID
            myTSDF_.TumbolName = .Name
            myTSDF_.ProvinceID = .ProvinceID
            myTSDF_.ProvinceName = .ProvinceName
        End If
    End With
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
    myTSDF_.Addr1 = Trim(txtAddr.Text)
End Sub

Private Sub txtCompanyName_Validate(Cancel As Boolean)
    myTSDF_.CompanyName = Trim(txtCompanyName.Text)
End Sub

Private Sub txtContactName_Validate(Cancel As Boolean)
    myTSDF_.ContactName = Trim(txtContactName.Text)
End Sub

Private Sub txtFaxNo_Validate(Cancel As Boolean)
    myTSDF_.FaxNo = Trim(txtFaxNo.Text)
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    myTSDF_.TelNo = Trim(txtTelNo.Text)
End Sub

Private Sub txtTSDFNo_Validate(Cancel As Boolean)
    myTSDF_.TSDFRegisNo = Trim(txtTSDFNo.Text)
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With myTSDF_
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
            With myTSDF_
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
            If myTSDF_.TumbolID = "" Then txtTumbol.Text = ""
        End If
    End If
End Sub

Private Sub txtZipCode_Validate(Cancel As Boolean)
    myTSDF_.ZipCode = Trim(txtZipCode.Text)
End Sub
