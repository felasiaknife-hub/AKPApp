VERSION 5.00
Begin VB.Form frmComRateConfig 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ตั้งค่า Config ในการคิดค่าคอมมิชชั่น - F016"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3540
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3540
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ค่าคอมสำหรับลูกค้าใหม่ Haz"
      ForeColor       =   &H80000008&
      Height          =   1245
      Index           =   3
      Left            =   3930
      TabIndex        =   34
      Top             =   2220
      Width           =   3795
      Begin VB.TextBox txtHzRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   9
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox txtHzPeriod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   10
         Top             =   750
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าคอมที่ให้"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   720
         TabIndex        =   38
         Top             =   420
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   2820
         TabIndex        =   37
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ระยะเวลาที่ให้"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   720
         TabIndex        =   36
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เดือน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   2820
         TabIndex        =   35
         Top             =   810
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ค่าคอมสำหรับลูกค้าใหม่ Non"
      ForeColor       =   &H80000008&
      Height          =   1245
      Index           =   2
      Left            =   90
      TabIndex        =   29
      Top             =   2220
      Width           =   3795
      Begin VB.TextBox txtNonPeriod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   8
         Top             =   750
         Width           =   945
      End
      Begin VB.TextBox txtNonRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   7
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เดือน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   2820
         TabIndex        =   33
         Top             =   810
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ระยะเวลาที่ให้"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   32
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   2820
         TabIndex        =   31
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าคอมที่ให้"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   30
         Top             =   420
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "โบนัสไตรมาส 3"
      ForeColor       =   &H80000008&
      Height          =   1605
      Index           =   1
      Left            =   3930
      TabIndex        =   22
      Top             =   540
      Width           =   3795
      Begin VB.TextBox txtTM3PC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   4
         Top             =   300
         Width           =   645
      End
      Begin VB.TextBox txtCRBonus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   5
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtManagerBonus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ต้องทำยอดขายได้"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   480
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   3120
         TabIndex        =   27
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานการตลาด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   26
         Top             =   750
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   25
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หัวหน้าส่วน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   24
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   3120
         TabIndex        =   23
         Top             =   1140
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Best Salers Bonus"
      ForeColor       =   &H80000008&
      Height          =   1605
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   540
      Width           =   3795
      Begin VB.TextBox txtBestSel3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1080
         Width           =   1245
      End
      Begin VB.TextBox txtBestSel2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   2
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtBestSel1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   1
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   3120
         TabIndex        =   21
         Top             =   1140
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รางวัลที่ 3"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1050
         TabIndex        =   20
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   19
         Top             =   750
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รางวัลที่ 2"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1050
         TabIndex        =   18
         Top             =   750
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   17
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รางวัลที่ 1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1050
         TabIndex        =   16
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.TextBox txtTopHz 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1380
      MaxLength       =   30
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2700
      TabIndex        =   14
      Top             =   180
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Top Haz Bonus"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   1110
   End
End
Attribute VB_Name = "frmComRateConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myComConfig_ As BWGComRateConfig

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
        If IsNumeric(Trim(txtTopHz.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtTopHz.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtBestSel1.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtBestSel1.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtBestSel2.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtBestSel2.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtBestSel3.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtBestSel3.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtTM3PC.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtTM3PC.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtCRBonus.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtCRBonus.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtManagerBonus.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtManagerBonus.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtNonRate.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtNonRate.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtNonPeriod.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtNonPeriod.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtHzRate.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtHzRate.SetFocus
            Exit Sub
        End If
        If IsNumeric(Trim(txtHzPeriod.Text)) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            txtHzPeriod.SetFocus
            Exit Sub
        End If
        If myComConfig_.isChanged Then
            myComConfig_.Save
        End If
        If myComConfig_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    sqlStr = "SELECT * FROM ComRateConfig"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Sub
     If tmpRec.EOF Then
        Set myComConfig_ = New BWGComRateConfig
        myComConfig_.ID = GetGUID
     Else
        Set myComConfig_ = New BWGComRateConfig
        myComConfig_.SetDataFromRecordset tmpRec
        ShowComConfig
     End If
     tmpRec.Close
     Set tmpRec = Nothing
     Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub ShowComConfig()
    If myComConfig_ Is Nothing Then Exit Sub
    With myComConfig_
        txtTopHz.Text = .TopHazBonus
        txtBestSel1.Text = .BestSellerBonus1
        txtBestSel2.Text = .BestSellerBonus2
        txtBestSel3.Text = .BestSellerBonus3
        txtTM3PC.Text = .Trimas3PC
        txtCRBonus.Text = .Trimas3CRBonus
        txtManagerBonus.Text = .Trimas3SupBonus
        txtNonRate.Text = .NewCustNonRate
        txtNonPeriod.Text = .NonMonthPeriod
        txtHzRate.Text = .NewCustHZRate
        txtHzPeriod.Text = .HzMonthPeriod
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myComConfig_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myComConfig_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtBestSel1_Validate(Cancel As Boolean)
    If IsNumeric(txtBestSel1.Text) Then
        myComConfig_.BestSellerBonus1 = txtBestSel1.Text
    Else
        myComConfig_.BestSellerBonus1 = 0
    End If
End Sub

Private Sub txtBestSel2_Validate(Cancel As Boolean)
    If IsNumeric(txtBestSel2.Text) Then
        myComConfig_.BestSellerBonus2 = txtBestSel2.Text
    Else
        myComConfig_.BestSellerBonus2 = 0
    End If
End Sub

Private Sub txtBestSel3_Validate(Cancel As Boolean)
    If IsNumeric(txtBestSel3.Text) Then
        myComConfig_.BestSellerBonus3 = txtBestSel3.Text
    Else
        myComConfig_.BestSellerBonus3 = 0
    End If
End Sub

Private Sub txtCRBonus_Validate(Cancel As Boolean)
    If IsNumeric(txtCRBonus.Text) Then
        myComConfig_.Trimas3CRBonus = txtCRBonus.Text
    Else
        myComConfig_.Trimas3CRBonus = 0
    End If
End Sub

Private Sub txtHzPeriod_Validate(Cancel As Boolean)
    If IsNumeric(txtHzPeriod.Text) Then
        myComConfig_.HzMonthPeriod = txtHzPeriod.Text
    Else
        myComConfig_.HzMonthPeriod = 0
    End If
End Sub

Private Sub txtHzRate_Validate(Cancel As Boolean)
    If IsNumeric(txtHzRate.Text) Then
        myComConfig_.NewCustHZRate = txtHzRate.Text
    Else
        myComConfig_.NewCustHZRate = 0
    End If
End Sub

Private Sub txtManagerBonus_Validate(Cancel As Boolean)
    If IsNumeric(txtManagerBonus.Text) Then
        myComConfig_.Trimas3SupBonus = txtManagerBonus.Text
    Else
        myComConfig_.Trimas3SupBonus = 0
    End If
End Sub

Private Sub txtNonPeriod_Validate(Cancel As Boolean)
    If IsNumeric(txtNonPeriod.Text) Then
        myComConfig_.NonMonthPeriod = txtNonPeriod.Text
    Else
        myComConfig_.NonMonthPeriod = 0
    End If
End Sub

Private Sub txtNonRate_Validate(Cancel As Boolean)
    If IsNumeric(txtNonRate.Text) Then
        myComConfig_.NewCustNonRate = txtNonRate.Text
    Else
        myComConfig_.NewCustNonRate = 0
    End If
End Sub

Private Sub txtTM3PC_Validate(Cancel As Boolean)
    If IsNumeric(txtTM3PC.Text) Then
        myComConfig_.Trimas3PC = txtTM3PC.Text
    Else
        myComConfig_.Trimas3PC = 0
    End If
End Sub

Private Sub txtTopHz_Validate(Cancel As Boolean)
    If IsNumeric(txtTopHz.Text) Then
        myComConfig_.TopHazBonus = txtTopHz.Text
    Else
        myComConfig_.TopHazBonus = 0
    End If
End Sub
