VERSION 5.00
Begin VB.Form frmTransportRate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "อัตราค่าขนส่ง - F126"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ราคาที่จ่ายให้กับ Vendor (รถพ่วง)"
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   3
      Left            =   150
      TabIndex        =   44
      Top             =   3420
      Width           =   7665
      Begin VB.TextBox txtVTrailerKM 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   930
         TabIndex        =   13
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtVTrailerTrip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3060
         TabIndex        =   14
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtVTrailerVol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5280
         TabIndex        =   15
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/กม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   2100
         TabIndex        =   47
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/เที่ยว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   4230
         TabIndex        =   46
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ลบ.ม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   6450
         TabIndex        =   45
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ราคาที่เรียกเก็บกับลูกค้า (รถพ่วง)"
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   2
      Left            =   150
      TabIndex        =   40
      Top             =   4260
      Width           =   7665
      Begin VB.TextBox txtCTrailerKM 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   930
         TabIndex        =   16
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtCTrailerTrip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3060
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtCTrailerVol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5280
         TabIndex        =   18
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/กม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   2100
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/เที่ยว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   4230
         TabIndex        =   42
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ลบ.ม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   6450
         TabIndex        =   41
         Top             =   330
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5250
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5250
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5250
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5250
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5250
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ราคาที่เรียกเก็บกับลูกค้า (รถเดี่ยว)"
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   2580
      Width           =   7665
      Begin VB.TextBox txtCustVolRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5280
         TabIndex        =   12
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txtCustTripRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3060
         TabIndex        =   11
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtCustKMRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   930
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ลบ.ม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   6450
         TabIndex        =   39
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/เที่ยว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4230
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/กม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   2100
         TabIndex        =   36
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ราคาที่จ่ายให้กับ Vendor (รถเดี่ยว)"
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   1740
      Width           =   7665
      Begin VB.TextBox txtVendorVolRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5280
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtVendorTripRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3060
         TabIndex        =   8
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtVendorKMRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   930
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ลบ.ม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   6450
         TabIndex        =   38
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/เที่ยว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   4230
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/กม."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   2100
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtDistance 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6270
      TabIndex        =   6
      Top             =   1320
      Width           =   765
   End
   Begin VB.TextBox txtToProvince 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   930
      Width           =   2565
   End
   Begin VB.TextBox txtToAumphur 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   930
      Width           =   2205
   End
   Begin VB.TextBox txtCarType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3270
      TabIndex        =   1
      Top             =   150
      Width           =   3765
   End
   Begin VB.TextBox txtSubID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   150
      Width           =   675
   End
   Begin VB.TextBox txtFromProvince 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   2565
   End
   Begin VB.TextBox txtFromAuphur 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   540
      Width           =   2205
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7980
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กม."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   7110
      TabIndex        =   31
      Top             =   1380
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ระยะทาง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   5610
      TabIndex        =   30
      Top             =   1380
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "จังหวัด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3900
      TabIndex        =   29
      Top             =   990
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอปลายทาง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   28
      Top             =   990
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรถ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2430
      TabIndex        =   27
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubID"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   1020
      TabIndex        =   26
      Top             =   210
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "จังหวัด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   3900
      TabIndex        =   25
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอต้นทาง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   540
      TabIndex        =   24
      Top             =   600
      Width           =   915
   End
End
Attribute VB_Name = "frmTransportRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myTransRate_ As BWGTransVendorChargeRate
Dim curFromAumphur_ As IMIAumphur
Dim curToAumphur_ As IMIAumphur
Dim curTruckType_ As BWGTruckType
Dim curAumphurType_ As String

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลอัตราค่าขนส่งรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myTransRate_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewTransRate
            txtSubID.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myTransRate_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myTransRate_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewTransRate
    txtSubID.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtSubID.Text) = "" Then
        MsgBox "กรุณาระบุ Sub ID", vbExclamation
        txtSubID.SetFocus
        Exit Sub
    End If

    If Trim(txtCarType.Text) = "" Then
        MsgBox "กรุณาใส่ประเภทรถ", vbExclamation
        txtCarType.SetFocus
        Exit Sub
    End If

    If Trim(txtFromAuphur.Text) = "" Then
        MsgBox "กรุณาระบุอำเภอต้นทางของการขนส่ง", vbExclamation
        txtFromAuphur.SetFocus
        Exit Sub
    End If

    If Trim(txtToAumphur.Text) = "" Then
        MsgBox "กรุณาระบุอำเภอปลายทางของการขนส่ง", vbExclamation
        txtToAumphur.SetFocus
        Exit Sub
    End If

    If Trim(txtDistance.Text) = "" Or IsNumeric(txtDistance.Text) = False Then
        MsgBox "กรุณาระบุระยะทางเป็นตัวเลขเท่านั้น", vbExclamation
        txtDistance.SetFocus
        Exit Sub
    End If

    If Trim(txtVendorKMRate.Text) = "" Or IsNumeric(txtVendorKMRate.Text) = False Then
        MsgBox "กรุณาระบุอัตราค่าขนส่งที่จ่ายให้กับ Vendor เป็นตัวเลขเท่านั้น", vbExclamation
        txtVendorKMRate.SetFocus
        Exit Sub
    End If
    If Trim(txtVendorTripRate.Text) = "" Or IsNumeric(txtVendorTripRate.Text) = False Then
        MsgBox "กรุณาระบุอัตราค่าขนส่งที่จ่ายให้กับ Vendor เป็นตัวเลขเท่านั้น", vbExclamation
        txtVendorTripRate.SetFocus
        Exit Sub
    End If

    If Trim(txtCustKMRate.Text) = "" Or IsNumeric(txtCustKMRate.Text) = False Then
        MsgBox "กรุณาระบุอัตราค่าขนส่งที่เรียกเก็บกับลูกค้าเป็นตัวเลขเท่านั้น", vbExclamation
        txtCustKMRate.SetFocus
        Exit Sub
    End If
    If Trim(txtCustTripRate.Text) = "" Or IsNumeric(txtCustTripRate.Text) = False Then
        MsgBox "กรุณาระบุอัตราค่าขนส่งที่เรียกเก็บกับลูกค้าเป็นตัวเลขเท่านั้น", vbExclamation
        txtCustTripRate.SetFocus
        Exit Sub
    End If

    If myTransRate_.isChanged Then
        myTransRate_.Save
    End If
    If myTransRate_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim myRes%
    If myTransRate_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            cmdSave_Click
            If myTransRate_.isChanged Then
                Exit Sub
            End If
        ElseIf myRes = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmTransportRateSearch
    With frmTransportRateSearch
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
    CreateNewTransRate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtSubID.Text = ""
    txtCarType.Text = ""
    txtFromAuphur.Text = ""
    txtFromProvince.Text = ""
    txtToAumphur.Text = ""
    txtToProvince.Text = ""
    txtDistance.Text = ""
    txtVendorKMRate.Text = ""
    txtVendorTripRate.Text = ""
    txtCustKMRate.Text = ""
    txtCustTripRate.Text = ""
    
    txtVendorVolRate.Text = ""
    txtCustVolRate.Text = ""
    txtVTrailerKM.Text = ""
    txtVTrailerTrip.Text = ""
    txtVTrailerVol.Text = ""
    txtCTrailerKM.Text = ""
    txtCTrailerTrip.Text = ""
    txtCTrailerVol.Text = ""
    
    Set curFromAumphur_ = Nothing
    Set curToAumphur_ = Nothing
    Set curTruckType_ = Nothing
    Set curFromAumphur_ = New IMIAumphur
    Set curToAumphur_ = New IMIAumphur
    Set curTruckType_ = New BWGTruckType
End Sub

Private Sub CreateNewTransRate()
    Set myTransRate_ = Nothing
    Set myTransRate_ = New BWGTransVendorChargeRate
    With myTransRate_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curFromAumphur_ Is Nothing Then
            .FromAumphurID = curFromAumphur_.ID
            .FromProvinceID = curFromAumphur_.ProvinceID
            .FromProvinceName = curFromAumphur_.ProvinceName
        End If
        If Not curToAumphur_ Is Nothing Then
            .ToAumphurID = curToAumphur_.ID
            .ToProvinceID = curToAumphur_.ProvinceID
            .ToProvinceName = curToAumphur_.ProvinceName
        End If
        If Not curTruckType_ Is Nothing Then
            .CarTypeID = curTruckType_.ID
            .CarTypeName = curTruckType_.TypeDesc
        End If
        .isChanged = False
    End With
End Sub

Public Sub SetDataTransportRate(selSite As BWGTransVendorChargeRate)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myTransRate_ = Nothing
    Set myTransRate_ = selSite
    If myTransRate_ Is Nothing Then Exit Sub
    With myTransRate_
        txtSubID.Text = .ChargeSubID
        txtDistance.Text = .EstDistance
        txtVendorKMRate.Text = .RatePerKM
        txtVendorTripRate = .RatePerTrip
        txtCustKMRate.Text = .CustRatePerKM
        txtCustTripRate.Text = .CustRatePerTrip
        
        txtVendorVolRate.Text = .RatePerVol
        txtCustVolRate.Text = .CustRatePerVol
        txtVTrailerKM.Text = .TrailerRateKM
        txtVTrailerTrip.Text = .TrailerRateTrip
        txtVTrailerVol.Text = .TrailerRateVol
        txtCTrailerKM.Text = .TrailerCustRateKM
        txtCTrailerTrip.Text = .TrailerCustRateTrip
        txtCTrailerVol.Text = .TrailerCustRateVol
        
        Set tmpCol = tmpS.TruckTypeSearch("TruckTypeID='" & .CarTypeID & "'")
        If tmpCol.Count > 0 Then SetDataTruckType tmpCol(1), False
        curAumphurType_ = "F"
        Set tmpCol = tmpS.AumphurSearch("AumphurID='" & .FromAumphurID & "'")
        If tmpCol.Count > 0 Then SetData tmpCol(1), False
        curAumphurType_ = "T"
        Set tmpCol = tmpS.AumphurSearch("AumphurID='" & .ToAumphurID & "'")
        If tmpCol.Count > 0 Then SetData tmpCol(1), False
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myTransRate_ Is Nothing Then Exit Sub
    If myTransRate_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myTransRate_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetData(selAumphur As IMIAumphur, Optional isAssigned As Boolean = True)
    If curAumphurType_ = "F" Then
        Set curFromAumphur_ = Nothing
        Set curFromAumphur_ = selAumphur
        If curFromAumphur_ Is Nothing Then Exit Sub
        With curFromAumphur_
            txtFromAuphur.Text = .Name
            txtFromProvince.Text = .ProvinceName
            If isAssigned Then
                myTransRate_.FromAumphurID = .ID
                myTransRate_.FromAumphurName = .Name
                myTransRate_.FromProvinceID = .ProvinceID
                myTransRate_.FromProvinceName = .ProvinceName
            End If
        End With
    ElseIf curAumphurType_ = "T" Then
        Set curToAumphur_ = Nothing
        Set curToAumphur_ = selAumphur
        If curToAumphur_ Is Nothing Then Exit Sub
        With curToAumphur_
            txtToAumphur.Text = .Name
            txtToProvince.Text = .ProvinceName
            If isAssigned Then
                myTransRate_.ToAumphurID = .ID
                myTransRate_.ToAumphurName = .Name
                myTransRate_.ToProvinceID = .ProvinceID
                myTransRate_.ToProvinceName = .ProvinceName
            End If
        End With
    End If
End Sub

Private Sub txtCarType_Validate(Cancel As Boolean)
    If curTruckType_ Is Nothing Then Set curTruckType_ = New BWGTruckType
    If Trim(txtCarType.Text) = "" And Trim(curTruckType_.TypeDesc) <> Trim(txtCarType.Text) Then
        Set curTruckType_ = Nothing
        Set curTruckType_ = New BWGTruckType
        With myTransRate_
            .CarTypeID = ""
            .CarTypeName = ""
        End With
    Else
        If Trim(curTruckType_.TypeDesc) <> Trim(txtCarType.Text) Then
            Set curTruckType_ = Nothing
            Set curTruckType_ = New BWGTruckType
            With myTransRate_
                .CarTypeID = ""
                .CarTypeName = ""
            End With
            Load frmTruckType
            With frmTruckType
                '.getCriteria (txtCarType.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myTransRate_.CarTypeID = "" Then txtCarType.Text = ""
        End If
    End If
End Sub

Private Sub txtCTrailerKM_Validate(Cancel As Boolean)
    If IsNumeric(txtCTrailerKM.Text) Then
        myTransRate_.TrailerCustRateKM = txtCTrailerKM.Text
    Else
        myTransRate_.TrailerCustRateKM = 0
    End If
End Sub

Private Sub txtCTrailerTrip_Validate(Cancel As Boolean)
    If IsNumeric(txtCTrailerTrip.Text) Then
        myTransRate_.TrailerCustRateTrip = txtCTrailerTrip.Text
    Else
        myTransRate_.TrailerCustRateTrip = 0
    End If
End Sub

Private Sub txtCTrailerVol_Validate(Cancel As Boolean)
    If IsNumeric(txtCTrailerVol.Text) Then
        myTransRate_.TrailerCustRateVol = txtCTrailerVol.Text
    Else
        myTransRate_.TrailerCustRateVol = 0
    End If
End Sub

Private Sub txtCustKMRate_Validate(Cancel As Boolean)
    If IsNumeric(txtCustKMRate.Text) Then
        myTransRate_.CustRatePerKM = Trim(txtCustKMRate.Text)
    Else
        myTransRate_.CustRatePerKM = 0
    End If
End Sub

Private Sub txtCustTripRate_Validate(Cancel As Boolean)
    If IsNumeric(txtCustTripRate.Text) Then
        myTransRate_.CustRatePerTrip = Trim(txtCustTripRate.Text)
    Else
        myTransRate_.CustRatePerTrip = 0
    End If
End Sub

Private Sub txtCustVolRate_Validate(Cancel As Boolean)
    If IsNumeric(txtCustVolRate.Text) Then
        myTransRate_.CustRatePerVol = txtCustVolRate.Text
    Else
        myTransRate_.CustRatePerVol = 0
    End If
End Sub

Private Sub txtDistance_Validate(Cancel As Boolean)
    If IsNumeric(txtDistance.Text) Then
        myTransRate_.EstDistance = Trim(txtDistance.Text)
    Else
        myTransRate_.EstDistance = 0
    End If
End Sub

Private Sub txtFromAuphur_GotFocus()
    curAumphurType_ = "F"
End Sub

Private Sub txtFromAuphur_Validate(Cancel As Boolean)
    If curFromAumphur_ Is Nothing Then Set curFromAumphur_ = New IMIAumphur
    If Trim(txtFromAuphur.Text) = "" And Trim(curFromAumphur_.Name) <> Trim(txtFromAuphur.Text) Then
        Set curFromAumphur_ = Nothing
        Set curFromAumphur_ = New IMIAumphur
        txtFromProvince.Text = ""
        With myTransRate_
            .FromAumphurID = ""
            .FromAumphurName = ""
            .FromProvinceID = ""
            .FromProvinceName = ""
        End With
    Else
        If Trim(curFromAumphur_.Name) <> Trim(txtFromAuphur.Text) Then
            Set curFromAumphur_ = Nothing
            Set curFromAumphur_ = New IMIAumphur
            txtFromProvince.Text = ""
            With myTransRate_
                .FromAumphurID = ""
                .FromAumphurName = ""
                .FromProvinceID = ""
                .FromProvinceName = ""
            End With
            Load frmAumphurSearch
            With frmAumphurSearch
                .getCriteria (txtFromAuphur.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myTransRate_.FromAumphurID = "" Then
                txtFromAuphur.Text = ""
                txtFromProvince.Text = ""
            End If
        End If
    End If
End Sub

Private Sub txtSubID_Validate(Cancel As Boolean)
    myTransRate_.ChargeSubID = Trim(txtSubID.Text)
End Sub

Private Sub txtToAumphur_GotFocus()
    curAumphurType_ = "T"
End Sub

Private Sub txtToAumphur_Validate(Cancel As Boolean)
    If curToAumphur_ Is Nothing Then Set curToAumphur_ = New IMIAumphur
    If Trim(txtToAumphur.Text) = "" And Trim(curToAumphur_.Name) <> Trim(txtToAumphur.Text) Then
        Set curToAumphur_ = Nothing
        Set curToAumphur_ = New IMIAumphur
        txtToProvince.Text = ""
        With myTransRate_
            .ToAumphurID = ""
            .ToAumphurName = ""
            .ToProvinceID = ""
            .ToProvinceName = ""
        End With
    Else
        If Trim(curToAumphur_.Name) <> Trim(txtToAumphur.Text) Then
            Set curToAumphur_ = Nothing
            Set curToAumphur_ = New IMIAumphur
            txtToProvince.Text = ""
            With myTransRate_
                .ToAumphurID = ""
                .ToAumphurName = ""
                .ToProvinceID = ""
                .ToProvinceName = ""
            End With
            Load frmAumphurSearch
            With frmAumphurSearch
                .getCriteria (txtToAumphur.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myTransRate_.ToAumphurID = "" Then
                txtToAumphur.Text = ""
                txtToProvince.Text = ""
            End If
        End If
    End If
End Sub

Public Sub SetDataTruckType(selSubType As BWGTruckType, Optional isAssigned As Boolean = True)
    Set curTruckType_ = Nothing
    Set curTruckType_ = selSubType
    If curTruckType_ Is Nothing Then Exit Sub
    With curTruckType_
        txtCarType.Text = .TypeDesc
        If isAssigned Then
            myTransRate_.CarTypeID = .ID
            myTransRate_.CarTypeName = .TypeDesc
        End If
    End With
End Sub

Private Sub txtVendorKMRate_Validate(Cancel As Boolean)
    If IsNumeric(txtVendorKMRate.Text) Then
        myTransRate_.RatePerKM = txtVendorKMRate.Text
    Else
        myTransRate_.RatePerKM = 0
    End If
End Sub

Private Sub txtVendorTripRate_Validate(Cancel As Boolean)
    If IsNumeric(txtVendorTripRate.Text) Then
        myTransRate_.RatePerTrip = txtVendorTripRate.Text
    Else
        myTransRate_.RatePerTrip = 0
    End If
End Sub

Private Sub txtVendorVolRate_Validate(Cancel As Boolean)
    If IsNumeric(txtVendorVolRate.Text) Then
        myTransRate_.RatePerVol = txtVendorVolRate.Text
    Else
        myTransRate_.RatePerVol = 0
    End If
End Sub

Private Sub txtVTrailerKM_Validate(Cancel As Boolean)
    If IsNumeric(txtVTrailerKM.Text) Then
        myTransRate_.TrailerRateKM = txtVTrailerKM.Text
    Else
        myTransRate_.TrailerRateKM = 0
    End If
End Sub

Private Sub txtVTrailerTrip_Validate(Cancel As Boolean)
    If IsNumeric(txtVTrailerTrip.Text) Then
        myTransRate_.TrailerRateTrip = txtVTrailerTrip.Text
    Else
        myTransRate_.TrailerRateTrip = 0
    End If
End Sub

Private Sub txtVTrailerVol_Validate(Cancel As Boolean)
    If IsNumeric(txtVTrailerVol.Text) Then
        myTransRate_.TrailerRateVol = txtVTrailerVol.Text
    Else
        myTransRate_.TrailerRateVol = 0
    End If
End Sub
