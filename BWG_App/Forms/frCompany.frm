VERSION 5.00
Begin VB.Form frmCompany 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F013"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtTsdf 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   19
      Top             =   5910
      Width           =   4425
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Height          =   1875
      Left            =   6030
      TabIndex        =   41
      Top             =   3960
      Width           =   3285
      Begin WasteManagment.ctlGrid grDisposeNo 
         Height          =   1695
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   2990
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000E&
      Caption         =   "ที่อยู่ (ภาษาอังกฤษ)"
      Height          =   1605
      Left            =   4710
      TabIndex        =   21
      Top             =   2340
      Width           =   4605
      Begin VB.TextBox txtAddrEN1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   10
         Top             =   330
         Width           =   4155
      End
      Begin VB.TextBox txtAddrEN2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   11
         Top             =   690
         Width           =   4155
      End
      Begin VB.TextBox txtAddrEN3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   12
         Top             =   1050
         Width           =   4155
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   60
      TabIndex        =   22
      Top             =   3960
      Width           =   5940
      Begin VB.TextBox txtCompanyRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   870
         TabIndex        =   17
         Top             =   1380
         Width           =   4935
      End
      Begin VB.TextBox txtWebsite 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   870
         TabIndex        =   16
         Top             =   1005
         Width           =   4935
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   870
         TabIndex        =   15
         Top             =   615
         Width           =   4935
      End
      Begin VB.TextBox txtTelNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   870
         TabIndex        =   13
         Top             =   225
         Width           =   2295
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3810
         TabIndex        =   14
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
         Left            =   135
         TabIndex        =   37
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
         Left            =   195
         TabIndex        =   35
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
         Left            =   315
         TabIndex        =   34
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
         Left            =   3285
         TabIndex        =   33
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
         Left            =   180
         TabIndex        =   32
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "ที่อยู่ (ภาษาไทย)"
      Height          =   1605
      Left            =   60
      TabIndex        =   20
      Top             =   2340
      Width           =   4605
      Begin VB.TextBox txtAddrTH3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   1050
         Width           =   4125
      End
      Begin VB.TextBox txtAddrTH2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   690
         Width           =   4125
      End
      Begin VB.TextBox txtAddrTH1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   210
         TabIndex        =   7
         Top             =   330
         Width           =   4125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "ข้อมูลบริษัท"
      Height          =   1470
      Left            =   30
      TabIndex        =   28
      Top             =   855
      Width           =   9315
      Begin VB.TextBox txtShortName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6840
         MaxLength       =   5
         TabIndex        =   5
         Top             =   990
         Width           =   915
      End
      Begin VB.TextBox txtDisposeCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8610
         MaxLength       =   100
         TabIndex        =   6
         Top             =   -150
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txtManagerName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6840
         MaxLength       =   100
         TabIndex        =   3
         Top             =   630
         Width           =   2205
      End
      Begin VB.TextBox txtCompanyCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   2205
      End
      Begin VB.TextBox txtTaxID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   4
         Top             =   990
         Width           =   1965
      End
      Begin VB.TextBox txtCompanyNameTH 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1890
         TabIndex        =   0
         Top             =   270
         Width           =   3945
      End
      Begin VB.TextBox txtCompanyNameEN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1890
         TabIndex        =   1
         Top             =   630
         Width           =   3945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อักษรย่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   6180
         TabIndex        =   42
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัวผู้กำจัดกากของเสีย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   6420
         TabIndex        =   40
         Top             =   -90
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อผู้จัดการ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   6030
         TabIndex        =   39
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสบริษัท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   6090
         TabIndex        =   38
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัวผู้เสียภาษี"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   225
         TabIndex        =   36
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อบริษัท (ภาษาไทย)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   31
         Top             =   345
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อบริษัท (ภาษาอังกฤษ)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   30
         Top             =   720
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   7215
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6405
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8265
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6405
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6405
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6405
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6405
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TSDF ที่เลือกใช้"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   43
      Top             =   5970
      Width           =   1125
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9375
      Y1              =   6315
      Y2              =   6315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูลบริษัท"
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
      TabIndex        =   29
      Top             =   0
      Width           =   1035
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
      Picture         =   "frCompany.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9345
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myCompany_ As BWGCompanyConfig
Dim curTSDF_ As BWGTSDF
Dim colCompanyGroup_ As Collection

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลกลุ่มย่อยของสินค้านี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myCompany_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewCompany
            txtCompanyNameTH.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myCompany_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCompany_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewCompany
    txtCompanyNameTH.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อบริษัทภาษาไทย", vbExclamation
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If

    If Trim(txtCompanyNameEN.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อบริษัทภาษาอังกฤษ", vbExclamation
        txtCompanyNameEN.SetFocus
        Exit Sub
    End If

    If Trim(txtCompanyCode.Text) = "" Then
        MsgBox "กรุณาใส่รหัสบริษัท", vbExclamation
        txtCompanyCode.SetFocus
        Exit Sub
    End If

    If myCompany_.isChanged Then
        myCompany_.Save
    End If
    If myCompany_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    If myCompany_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCompany_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCompanySearch
    With frmCompanySearch
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
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With grDisposeNo.getGridObj
        .Cols = 2
        .ColWidth(0) = 0
        .ColWidth(1) = 2895
        .Rows = 20
        .TextMatrix(0, 1) = "เลขประจำตัวผู้กำจัดกาก"
        .ColAlignment(1) = 1
        .FixedAlignment(1) = 1
    End With
    CreateNewCompany
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grDisposeNo" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtCompanyCode.Text = ""
    txtCompanyNameTH.Text = ""
    txtDisposeCode.Text = ""
    txtCompanyNameEN.Text = ""
    txtShortName.Text = ""
    txtManagerName.Text = ""
    txtAddrTH1.Text = ""
    txtAddrTH2.Text = ""
    txtAddrTH3.Text = ""
    txtAddrEN1.Text = ""
    txtAddrEN2.Text = ""
    txtAddrEN3.Text = ""
    txtTaxID.Text = ""
    txtTelNo.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtWebsite.Text = ""
    txtCompanyRem.Text = ""
    grDisposeNo.ClearAllData
    grDisposeNo.getGridObj.Rows = 20
    txtTsdf.Text = ""
    Set curTSDF_ = Nothing
    Set curTSDF_ = New BWGTSDF
End Sub

Private Sub CreateNewCompany()
    Set myCompany_ = Nothing
    Set myCompany_ = New BWGCompanyConfig
    With myCompany_
        .ID = GetGUID ' getNewProductID("CC")
        .isChanged = False
    End With
End Sub

Public Sub SetDataCompany(selCompany As BWGCompanyConfig)
Dim x%, tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myCompany_ = Nothing
    Set myCompany_ = selCompany
    If myCompany_ Is Nothing Then Exit Sub
    With myCompany_
        Set tmpCol = tmpS.TSDFSearch("TSDFID='" & .TSDFID & "'")
        If tmpCol.Count > 0 Then
            SetDataTSDF tmpCol(1)
        Else
            txtTsdf.Text = ""
            Set curTSDF_ = Nothing
            Set curTSDF_ = New BWGTSDF
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
        txtCompanyCode.Text = .CompanyCode
        txtCompanyNameTH.Text = .CompanyNameTH
        txtCompanyNameEN.Text = .CompanyNameEN
        txtShortName.Text = .ShortName
        txtManagerName.Text = .ManagerName
        txtAddrTH1.Text = .AddrTH1
        txtAddrTH2.Text = .AddrTH2
        txtAddrTH3.Text = .AddrTH3
        txtAddrEN1.Text = .AddrEN1
        txtAddrEN2.Text = .AddrEN2
        txtAddrEN3.Text = .AddrEN3
        txtDisposeCode.Text = .DisposerNo
        txtTaxID.Text = .TaxID
        txtTelNo.Text = .PhoneNo
        txtFax.Text = .FaxNo
        txtEmail.Text = .EmailAddr
        txtWebsite.Text = .Website
        txtCompanyRem.Text = .CompanyRemark
    End With
    grDisposeNo.ClearAllData
    With grDisposeNo.getGridObj
        Set tmpCol = myCompany_.DisposerNoList
        .Rows = 20 + tmpCol.Count
        For x = 1 To tmpCol.Count
            .TextMatrix(x, 0) = tmpCol(x).ID
            .TextMatrix(x, 1) = tmpCol(x).DisposerNo
        Next
        Set tmpCol = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myCompany_ Is Nothing Then Exit Sub
    If myCompany_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCompany_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub grDisposeNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpKey As String
    With grDisposeNo.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            If MsgBox("ท่านต้องการลบเลขที่ผู้กำจัดของเสีย '" & .TextMatrix(.row, 1) & "' หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                If isExist(myCompany_.DisposerNoList, tmpKey) Then
                    myCompany_.DisposerNoList.Remove tmpKey
                End If
                .RemoveItem .row
                myCompany_.isChanged = True
            End If
        End If
    End With
End Sub

Private Sub grDisposeNo_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGDisposerNo, tmpKey As String
    tmpKey = "'" & Trim(grDisposeNo.getGridObj.TextMatrix(row, 0)) & "'"
    If isExist(myCompany_.DisposerNoList, tmpKey) Then
        Set tmpDet = myCompany_.DisposerNoList(tmpKey)
    Else
        Set tmpDet = New BWGDisposerNo
        tmpDet.ID = GetGUID
        grDisposeNo.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        tmpDet.companyID = myCompany_.ID
        myCompany_.DisposerNoList.Add tmpDet, "'" & tmpDet.ID & "'"
    End If
    If col = 1 Then
        tmpDet.DisposerNo = Trim(grDisposeNo.getGridObj.TextMatrix(row, 1))
        myCompany_.isChanged = True
    End If
    Set tmpDet = Nothing
End Sub

Private Sub txtAddrEN1_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.AddrEN1 = Trim(txtAddrEN1.Text)
End Sub

Private Sub txtAddrEN2_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.AddrEN2 = Trim(txtAddrEN2.Text)
End Sub

Private Sub txtAddrEN3_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.AddrEN3 = Trim(txtAddrEN3.Text)
End Sub

Private Sub txtAddrTH1_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.AddrTH1 = Trim(txtAddrTH1.Text)
End Sub

Private Sub txtAddrTH2_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.AddrTH2 = Trim(txtAddrTH2.Text)
End Sub

Private Sub txtAddrTH3_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.AddrTH3 = Trim(txtAddrTH3.Text)
End Sub

Private Sub txtCompanyCode_Validate(Cancel As Boolean)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    If Trim(txtCompanyCode.Text) <> "" And Trim(txtCompanyCode.Text) <> Trim(myCompany_.CompanyCode) Then
        Set tmpCol = tmpS.CompanySearch("CompanyCode='" & Trim(txtCompanyCode.Text) & "'")
        If tmpCol.Count > 0 Then
            If Trim(tmpCol(1).ID) <> Trim(myCompany_.ID) Then
                MsgBox "รหัสบริษัทนี้ถูกกำหนดให้กับบริษัทอื่นแล้ว กรุณาเปลี่ยนรหัสบริษัทใหม่", vbExclamation
                Cancel = True
                Set tmpCol = Nothing
                Set tmpS = Nothing
                Exit Sub
            End If
        End If
    End If
    myCompany_.CompanyCode = Trim(txtCompanyCode.Text)
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub txtCompanyNameEN_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.CompanyNameEN = Trim(txtCompanyNameEN.Text)
End Sub

Private Sub txtCompanyNameTH_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.CompanyNameTH = Trim(txtCompanyNameTH.Text)
End Sub

Private Sub txtCompanyRem_Change()
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.CompanyRemark = Trim(txtCompanyRem.Text)
End Sub

Private Sub txtDisposeCode_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.DisposerNo = Trim(txtDisposeCode.Text)
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.EmailAddr = Trim(txtEmail.Text)
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.FaxNo = Trim(txtFax.Text)
End Sub

Private Sub txtManagerName_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.ManagerName = Trim(txtManagerName.Text)
End Sub

Private Sub txtShortName_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.ShortName = Trim(txtShortName.Text)
End Sub

Private Sub txtTaxID_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.TaxID = Trim(txtTaxID.Text)
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.PhoneNo = Trim(txtTelNo.Text)
End Sub

Private Sub txtTsdf_Validate(Cancel As Boolean)
    If curTSDF_ Is Nothing Then Set curTSDF_ = New BWGTSDF
    If Trim(txtTsdf.Text) = "" And Trim(curTSDF_.CompanyName) <> Trim(txtTsdf.Text) Then
        Set curTSDF_ = Nothing
        Set curTSDF_ = New BWGTSDF
        With myCompany_
            .TSDFID = ""
            .TSDFName = ""
        End With
    Else
        If Trim(curTSDF_.CompanyName) <> Trim(txtTsdf.Text) Then
            Set curTSDF_ = Nothing
            Set curTSDF_ = New BWGTSDF
            With myCompany_
                .TSDFID = ""
                .TSDFName = ""
            End With
            Load frmTSDFSearch
            With frmTSDFSearch
                .getCriteria (txtTsdf.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myCompany_.TSDFID = "" Then txtTsdf.Text = ""
        End If
    End If
End Sub

Private Sub txtWebSite_Validate(Cancel As Boolean)
    If myCompany_ Is Nothing Then
        CreateNewCompany
    End If
    myCompany_.Website = Trim(txtWebsite.Text)
End Sub

Public Sub SetDataTSDF(selTsdf As BWGTSDF, Optional isAssigned As Boolean = True)
    Set curTSDF_ = selTsdf
    With curTSDF_
        txtTsdf.Text = .CompanyName
        If isAssigned Then
            myCompany_.TSDFID = .ID
            myCompany_.TSDFName = .CompanyName
        End If
    End With
End Sub
