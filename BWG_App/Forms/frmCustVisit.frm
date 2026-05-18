VERSION 5.00
Begin VB.Form frmCustVisit 
   Caption         =   "ลูกค้าเข้าเยี่ยมชมศูนย์"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   13200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkOnline 
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6525
      MaskColor       =   &H00C00000&
      TabIndex        =   58
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CheckBox ChkNoCome 
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6525
      MaskColor       =   &H000000C0&
      TabIndex        =   46
      Top             =   7650
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CheckBox ChkInform 
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4590
      MaskColor       =   &H000000C0&
      TabIndex        =   38
      Top             =   8190
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CheckBox ChkUse 
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   4590
      MaskColor       =   &H00008000&
      TabIndex        =   34
      Top             =   7920
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.TextBox TxtPoint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   420
      Left            =   2790
      TabIndex        =   32
      Text            =   "0"
      Top             =   7650
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   840
      Left            =   11250
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7650
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   840
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7650
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   840
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7650
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   840
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7650
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "ข้อมูลลูกค้า"
      Height          =   6945
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   13245
      Begin VB.TextBox txtNameDisplay 
         Height          =   330
         Left            =   4770
         TabIndex        =   59
         Top             =   225
         Width           =   2175
      End
      Begin VB.CheckBox ChkCheckList 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   12060
         MaskColor       =   &H0000FF00&
         TabIndex        =   54
         Top             =   1035
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.Frame Frame4 
         Caption         =   "ข้อเสนอแนะ"
         Height          =   870
         Left            =   0
         TabIndex        =   48
         Top             =   1710
         Width           =   13245
         Begin VB.CheckBox ChkCCTV 
            Caption         =   "ภาพจากการดูกล้อง CCTV"
            Height          =   195
            Left            =   135
            TabIndex        =   56
            Top             =   540
            Width           =   2400
         End
         Begin VB.CheckBox ChkWelcome 
            Caption         =   "การต้อนรับของเจ้าหน้าที่และการตอบข้อซักถาม"
            Height          =   195
            Left            =   135
            TabIndex        =   53
            Top             =   270
            Width           =   3660
         End
         Begin VB.CheckBox ChkProcess 
            Caption         =   "ความเหมาะสมของระยะเวลาในการดูทำลายสินค้า"
            Height          =   195
            Left            =   4410
            TabIndex        =   52
            Top             =   270
            Width           =   3885
         End
         Begin VB.CheckBox ChkReady 
            Caption         =   "ความเหมาะสมของสถานที่ตรวจนับสินค้า"
            Height          =   195
            Left            =   8820
            TabIndex        =   51
            Top             =   270
            Width           =   3795
         End
         Begin VB.CheckBox ChkClean 
            Caption         =   "ความสะอาดของพื้นที่โดยรอบ"
            Height          =   195
            Left            =   4410
            TabIndex        =   50
            Top             =   540
            Width           =   2400
         End
         Begin VB.CheckBox ChkPPE 
            Caption         =   "การสวมใส่อุปกรณ์ PPE ของพนักงาน"
            Height          =   195
            Left            =   8820
            TabIndex        =   49
            Top             =   540
            Width           =   2895
         End
      End
      Begin VB.CheckBox ChkNewCust 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   10755
         MaskColor       =   &H0000FF00&
         TabIndex        =   39
         Top             =   1035
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.TextBox txtSale2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8550
         TabIndex        =   35
         Top             =   990
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "รายชื่อลูกค้า"
         Height          =   4305
         Left            =   45
         TabIndex        =   23
         Top             =   2610
         Width           =   6495
         Begin WasteManagment.ctlGrid grCustomer 
            Height          =   4050
            Left            =   45
            TabIndex        =   29
            Top             =   225
            Width           =   6390
            _extentx        =   11033
            _extenty        =   7144
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "เจ้าหน้าที่ดูแลลูกค้า"
         Height          =   4305
         Left            =   6705
         TabIndex        =   24
         Top             =   2610
         Width           =   6495
         Begin WasteManagment.ctlGrid grStaff 
            Height          =   4050
            Left            =   45
            TabIndex        =   30
            Top             =   225
            Width           =   6390
            _extentx        =   11033
            _extenty        =   7144
         End
      End
      Begin VB.TextBox TxtRemark 
         Height          =   330
         Left            =   1170
         TabIndex        =   22
         Top             =   1350
         Width           =   11895
      End
      Begin VB.TextBox TxtTime 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   11160
         TabIndex        =   3
         Top             =   255
         Width           =   1905
      End
      Begin VB.ComboBox CboType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCustVisit.frx":0000
         Left            =   4755
         List            =   "frmCustVisit.frx":0010
         TabIndex        =   18
         Top             =   990
         Width           =   2160
      End
      Begin VB.TextBox TxtTotal 
         Height          =   330
         Left            =   1170
         TabIndex        =   16
         Text            =   "1"
         Top             =   990
         Width           =   780
      End
      Begin VB.TextBox TxtMail 
         Height          =   330
         Left            =   10485
         TabIndex        =   14
         Top             =   630
         Width           =   2580
      End
      Begin VB.TextBox TxtTel 
         Height          =   330
         Left            =   7740
         TabIndex        =   12
         Top             =   630
         Width           =   1770
      End
      Begin VB.TextBox TxtPosition 
         Height          =   330
         Left            =   4770
         TabIndex        =   10
         Top             =   630
         Width           =   2175
      End
      Begin VB.TextBox TxtContact 
         Height          =   330
         Left            =   1170
         TabIndex        =   8
         Top             =   630
         Width           =   2400
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Top             =   255
         Width           =   2400
      End
      Begin WasteManagment.ctlDate dtWorkDate 
         Height          =   315
         Left            =   8370
         TabIndex        =   2
         Top             =   255
         Width           =   1785
         _extentx        =   3149
         _extenty        =   556
         backcolor       =   -2147483633
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "ซื่อที่แสดง"
         Height          =   195
         Left            =   3720
         TabIndex        =   60
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "เช็คลิสต์"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   12375
         TabIndex        =   55
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ลูกค้าใหม่"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   11070
         TabIndex        =   43
         Top             =   1035
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "เซลล์ผู้รับผิดชอบ"
         Height          =   195
         Left            =   7335
         TabIndex        =   36
         Top             =   1035
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "หมายเหตุ"
         Height          =   195
         Left            =   405
         TabIndex        =   21
         Top             =   1395
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "เวลา"
         Height          =   195
         Left            =   10710
         TabIndex        =   20
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วัตถุประสงค์ในการเข้าศูนย์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   2745
         TabIndex        =   19
         Top             =   1035
         Width           =   1905
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "คน"
         Height          =   195
         Left            =   2025
         TabIndex        =   17
         Top             =   1035
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "จำนวน"
         Height          =   195
         Left            =   585
         TabIndex        =   15
         Top             =   1035
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "E-mail"
         Height          =   195
         Left            =   9990
         TabIndex        =   13
         Top             =   675
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "โทร"
         Height          =   195
         Left            =   7425
         TabIndex        =   11
         Top             =   675
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ตำแหน่ง"
         Height          =   195
         Left            =   4095
         TabIndex        =   9
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ผู้ติดต่อ"
         Height          =   195
         Left            =   585
         TabIndex        =   7
         Top             =   675
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า บริษัท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   5
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่เข้าศูนย์"
         Height          =   195
         Index           =   0
         Left            =   7395
         TabIndex        =   4
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.CheckBox ChkIsComplete 
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4590
      MaskColor       =   &H00C00000&
      TabIndex        =   33
      Top             =   7650
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ลูกค้าดูผ่านออนไลน์"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6795
      TabIndex        =   57
      Top             =   7920
      Width           =   1380
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "แจ้งว่ามาแต่ลูกค้าไม่มา"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   6795
      TabIndex        =   47
      Top             =   7650
      Width           =   1605
   End
   Begin VB.Label lblAverage 
      AutoSize        =   -1  'True
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2790
      TabIndex        =   45
      Top             =   8100
      Width           =   240
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "ค่าเฉลี่ย :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1755
      TabIndex        =   44
      Top             =   8100
      Width           =   960
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ไม่ได้แจ้งล่วงหน้า"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4860
      TabIndex        =   42
      Top             =   8190
      Width           =   1230
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ตัดสินใจใช้บริการ"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   4860
      TabIndex        =   41
      Top             =   7920
      Width           =   1260
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "เสร็จสิ้นแล้ว"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4860
      TabIndex        =   40
      Top             =   7650
      Width           =   855
   End
   Begin VB.Label Label13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   37
      Top             =   7695
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "คะแนนประเมินจากลูกค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   90
      TabIndex        =   31
      Top             =   7695
      Width           =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   7560
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้าประสงค์ขอเข้าศูนย์"
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
      Left            =   150
      TabIndex        =   6
      Top             =   90
      Width           =   2145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmCustVisit.frx":0058
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15720
   End
End
Attribute VB_Name = "frmCustVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myJobOpen_ As BWGJobData
Dim myCustVisit_ As BWGCustVisitor
Dim curCustomer_ As BWGCustomer
Dim isDetailChange_ As Boolean
Dim curCar_ As BWGJobDataCarUsed
Dim Index As Integer
Public gVal As String
Public isNew As Boolean
Dim curSaleStaff_ As BWGCompanyStaff

Private Sub CboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub CboType_Validate(Cancel As Boolean)
    If cboType.ListIndex = 1 Then
        Label13.Caption = "/60"
    ElseIf cboType.ListIndex = 2 Then
        Label13.Caption = "/35"
    ElseIf cboType.ListIndex = 3 Then
        Label13.Caption = "/55"
    Else
        Label13.Caption = ""
    End If
    Label13.Caption = "100%"
    myCustVisit_.VisitType = Trim(cboType.Text)
End Sub

Private Sub ChkCCTV_Click()
    If ChkCCTV.Value = 1 Then
        myCustVisit_.CCTV = True
    Else
        myCustVisit_.CCTV = False
    End If
End Sub

Private Sub ChkCheckList_Click()
    If ChkCheckList.Value = 1 Then
        myCustVisit_.CheckList = True
    Else
        myCustVisit_.CheckList = False
    End If
End Sub

Private Sub ChkClean_Click()
    If ChkClean.Value = 1 Then
        myCustVisit_.FieldClean = True
    Else
        myCustVisit_.FieldClean = False
    End If
End Sub

Private Sub ChkInform_Click()
    If ChkInform.Value = 1 Then
        myCustVisit_.IsInform = True
    Else
        myCustVisit_.IsInform = False
    End If
End Sub

Private Sub ChkIsComplete_Click()
    If ChkIsComplete.Value = 1 Then
        myCustVisit_.isComplete = True
    Else
        myCustVisit_.isComplete = False
    End If
End Sub

Private Sub ChkNewCust_Click()
    If ChkNewCust.Value = 1 Then
        myCustVisit_.IsNewCust = True
    Else
        myCustVisit_.IsNewCust = False
    End If
End Sub

Private Sub ChkNoCome_Click()
    If ChkNoCome.Value = 1 Then
        myCustVisit_.IsNoCome = True
    Else
        myCustVisit_.IsNoCome = False
    End If
End Sub

Private Sub ChkOnline_Click()
    If ChkOnline.Value = 1 Then
        myCustVisit_.LookOnline = True
    Else
        myCustVisit_.LookOnline = False
    End If

End Sub

Private Sub ChkPPE_Click()
    If ChkPPE.Value = 1 Then
        myCustVisit_.EmpPPE = True
    Else
        myCustVisit_.EmpPPE = False
    End If
End Sub

Private Sub ChkProcess_Click()
    If ChkProcess.Value = 1 Then
        myCustVisit_.WorkProcess = True
    Else
        myCustVisit_.WorkProcess = False
    End If
End Sub

Private Sub ChkReady_Click()
    If ChkReady.Value = 1 Then
        myCustVisit_.EmpReady = True
    Else
        myCustVisit_.EmpReady = False
    End If
End Sub

Private Sub ChkUse_Click()
    If ChkUse.Value = 1 Then
        myCustVisit_.isUse = True
    Else
        myCustVisit_.isUse = False
    End If
End Sub

Private Sub ChkWelcome_Click()
    If ChkWelcome.Value = 1 Then
        myCustVisit_.Welcome = True
    Else
        myCustVisit_.Welcome = False
    End If

End Sub

Private Sub cmdCancel_Click()
    Dim result
    result = MsgBox("ต้องการยกเลิกกำหนดการเยี่ยมชมนี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        DBConnExc "Update tbCustomerVisitHeader Set IsCancel = 'Y' WHERE VisitID='" & myCustVisit_.id & "' "
    End If
    Unload Me
End Sub

Private Sub cmdClear_Click()
    ClearScreen
    CreateNewVisit
    dtWorkDate.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CreateNewVisit()
    Set myCustVisit_ = Nothing
    Set myCustVisit_ = New BWGCustVisitor
    With myCustVisit_
        .id = GetGUID
    End With
    txtCustomer.Enabled = True
End Sub

Private Sub CmdSave_Click()
Dim i As Integer
Dim x As Integer
    If myCustVisit_.CustomerID = "" Then
        MsgBox "กรุณาระบุลูกค้าก่อน", vbExclamation + vbOKOnly, "ระบุข้อมูลไม่ครบถ้วน"
        txtCustomer.Text = ""
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboType.ListIndex = 0 Then
        MsgBox "กรุณาระบุวัตถุประสงค์ในการเข้าศูนย์", vbExclamation + vbOKOnly, "ข้อมูลสำคัญ"
        Exit Sub
    End If
    If Val(TxtTotal.Text) = 0 Then
        MsgBox "กรุณาระบุจำนวนผู้เข้าศูนย์ อย่างน้อย 1 คน", vbExclamation + vbOKOnly, "ข้อมูลสำคัญ"
        TxtTotal.SetFocus
        Exit Sub
    End If
    If grCustomer.getGridObj.TextMatrix(1, 1) = "" Then
        MsgBox "กรุณาระบุชื่อผู้เข้าเยี่ยมชมศูนย์ อย่างน้อย 1 คน", vbExclamation + vbOKOnly, "ข้อมูลสำคัญ"
        Exit Sub
    End If
    If grStaff.getGridObj.TextMatrix(1, 1) = "" Then
        MsgBox "กรุณาระบุพนักงานรับลูกค้า อย่างน้อย 1 คน", vbExclamation + vbOKOnly, "ข้อมูลสำคัญ"
        Exit Sub
    End If
        If cboType.Text = "เยี่ยมชมศูนย์" Then
            myCustVisit_.FullPoint = "60"
        ElseIf cboType.Text = "ตรวจประเมินประจำปี" Then
            myCustVisit_.FullPoint = "30"
        ElseIf cboType.Text = "ลูกค้าติดตามดูทำลาย" Then
            myCustVisit_.FullPoint = "55"
        End If

                DBConnExc "delete tbCustomerVisitDetail WHERE VisitID='" & myCustVisit_.id & " ' "
                DBConnExc "delete tbReceptionStaff WHERE VisitID='" & myCustVisit_.id & " ' "
    With grCustomer.getGridObj
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                DBConnExc "Insert into tbCustomerVisitDetail(ID, VisitID, VisitorName,Position) Values ('" & GetGUID & "','" & myCustVisit_.id & "','" & .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "')"
            End If
        Next i
    End With
    With grStaff.getGridObj
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                DBConnExc "Insert into tbReceptionStaff(VisitID, StaffName,StaffID) Values ('" & myCustVisit_.id & "','" & .TextMatrix(i, 1) & "'," & .TextMatrix(i, 3) & ")"
            End If
        Next i
    End With
        myCustVisit_.VisitDate = dtWorkDate.ValueYMD
        myCustVisit_.Save
        If myCustVisit_.isChanged = False Then
            MsgBox "เก็บข้อมูลเรียบร้อย", vbInformation + vbOKOnly, "บันทึกข้อมูล"
        End If
        Unload Me
End Sub

Private Sub dtWorkDate_Validate(Cancel As Boolean)
    myCustVisit_.VisitDate = dtWorkDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "grCustomer" And Me.ActiveControl.Name <> "grStaff" Then SendKeys "{Tab}"
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    CreateNewVisit
    dtWorkDate.ValueYMD = TodayDate
    SetGrid
    PopulateReception
    cboType.ListIndex = 0
'    If LCase(CurrentUser.StaffCode) = "pr" Or LCase(CurrentUser.StaffCode) = "env" Or LCase(CurrentUser.StaffCode) = "akpenvi" _
'    Or LCase(CurrentUser.StaffCode) = "admin" Or LCase(CurrentUser.StaffCode) = "env02" Then
'    cmdSave.Visible = True
'    cmdClear.Visible = True
'    cmdCancel.Visible = True
'    Else
'    cmdSave.Visible = False
'    cmdClear.Visible = False
'    cmdCancel.Visible = False
'    End If
ChkIsComplete.ForeColor = vbBlue
ChkUse.ForeColor = vbGreen
ChkInform.ForeColor = vbRed
ChkNewCust.ForeColor = vbGreen
End Sub
Private Sub SetGrid()
Dim x%
    With grCustomer.getGridObj
        .Rows = 100
        .Cols = 3
        .TextMatrix(0, 1) = "ชื่อ-สกุล"
        .TextMatrix(0, 2) = "ตำแหน่ง"
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 3480
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    With grStaff.getGridObj
        .Rows = 10
        .Cols = 4
        grStaff.ColInputType(1) = ComboBoxEnum
        .TextMatrix(0, 1) = "เจ้าหน้าที่รับลูกค้า"
        .TextMatrix(0, 2) = "แผนก"
        .TextMatrix(0, 3) = "StaffID"
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(2) = 3480
        .ColWidth(3) = 0
        
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With

End Sub

Private Sub grCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grCustomer.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบชื่อผู้เยี่ยมชม " & .TextMatrix(.row, 1) & " หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                DBConnExc "delete tbCustomerVisitDetail WHERE ID=" & tmpKey & " "
                .TextMatrix(.row, 1) = ""
                .RemoveItem .row
            End If
        ElseIf KeyCode = 46 And Trim(.TextMatrix(.row, 0)) = "" Then
            selResult = MsgBox("ต้องการลบชื่อผู้เยี่ยมชม " & .TextMatrix(.row, 1) & " หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                .RemoveItem .row
            End If
        End If
        .Rows = .Rows + 1
    End With
End Sub

Private Sub grStaff_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grStaff.getGridObj
        If Trim(.TextMatrix(2, 1)) = "" Then
            MsgBox "ต้องมีพนักงานอย่างน้อย 1 คน เพื่อรับลูกค้า", vbExclamation + vbOKOnly, "ผิดพลาด"
            Exit Sub
        End If
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบชื่อพนักงานรับลูกค้า " & .TextMatrix(.row, 1) & " หรือไม่?" & vbCrLf & _
            "ข้อมูลจะถูกลบทันทีหลังจากกดปุ่ม Yes", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                DBConnExc "delete tbReceptionStaff WHERE ID=" & tmpKey & "  "
                grStaff.ClearAllData
                SetGrid
                ShowRecStaff
            End If
        ElseIf KeyCode = 46 And Trim(.TextMatrix(.row, 0)) = "" Then
            selResult = MsgBox("ต้องการลบชื่อพนักงานรับลูกค้า " & .TextMatrix(.row, 1) & " หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                grStaff.ClearAllData
                SetGrid
                ShowRecStaff
            End If
        End If
        .Rows = .Rows + 1
    End With

End Sub

Private Sub grStaff_RolColChanged(row As Integer, col As Integer)
Dim x%
    With grStaff.getGridObj
            If col = 1 Then
                With grStaff.getColObject(1)
                    .Clear
                        For x = 1 To colReception_.Count
                            .AddItem colReception_(x).StaffName
                        Next
                End With
            End If
    End With
End Sub

Private Sub grStaff_UpdateValue(row As Integer, col As Integer)
    With grStaff.getGridObj
        Select Case col
            Case 1
                Dim tSql As String
                Dim rs As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM tbReception " & vbCrLf
                    tSql = tSql & "where staffname= '" & Trim(.TextMatrix(row, 1)) & "' " & vbCrLf
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    If rs.RecordCount > 0 Then
                        .TextMatrix(row, 2) = rs!Dept
                        .TextMatrix(row, 3) = rs!id
                    End If
        End Select
    End With
End Sub

Private Sub TxtContact_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub TxtContact_Validate(Cancel As Boolean)
    myCustVisit_.ContactName = Trim(TxtContact.Text)
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
Dim tmpCustID As String
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myCustVisit_
            .CustomerID = ""
            .CompanyName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.id <> tmpCustID Then
                grCustomer.ClearAllData
                grStaff.ClearAllData
            End If
        End If
    End If
                SetGrid
End Sub
Sub ClearScreen()
    txtCustomer.Text = ""
    Set curCustomer_ = Nothing
    TxtContact.Text = ""
    TxtPosition.Text = ""
    TxtTel.Text = ""
    TxtMail.Text = ""
    TxtRemark.Text = ""
    TxtTime.Text = ""
    TxtTotal.Text = ""
    txtNameDisplay.Text = ""
    dtWorkDate.ValueYMD = TodayDate
    grCustomer.ClearAllData
    grStaff.ClearAllData
    SetGrid
End Sub
Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
On Error GoTo ErrD
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
        myCustVisit_.CompanyName = .CustomerName
        myCustVisit_.CustomerID = .id
        myCustVisit_.CustomerCode = .CustomerCode
        myCustVisit_.CustTypeFlag = .CustomerType
        myCustVisit_.NameDisplay = .CustomerName
        txtNameDisplay.Text = .CustomerName
    Dim tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
        tSql = "SELECT SaleStaffName FROM vw_CustomerSearch " & vbCrLf
        tSql = tSql & "where CustomerID= '" & myCustVisit_.CustomerID & "' " & vbCrLf
        Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
        If rs.RecordCount > 0 Then
            txtSale2.Text = rs!SaleStaffName
            myCustVisit_.SaleStaff = rs!SaleStaffName
        End If

    End With
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub
Public Sub SetDataCustVisit(selVisit As BWGCustVisitor)
Dim tmpStr, i As Integer
Dim tmpCol As Collection
Dim tPoint As Integer
Dim Ave As Double
     Set myCustVisit_ = selVisit
    With myCustVisit_
        txtCustomer.Text = .CompanyName
        txtNameDisplay.Text = .NameDisplay
        dtWorkDate.ValueYMD = .VisitDate
        TxtTime.Text = .VisitTime
        TxtContact.Text = .ContactName
        TxtPosition.Text = .Position
        TxtTel.Text = .Tel
        TxtMail.Text = .Email
        TxtTotal.Text = .TotalPeople
        If .VisitType = "เยี่ยมชมศูนย์" Then
            cboType.ListIndex = 1
            Label13.Caption = "/60"
        ElseIf .VisitType = "ตรวจประเมินประจำปี" Then
            cboType.ListIndex = 2
            Label13.Caption = "/35"
        ElseIf .VisitType = "ลูกค้าติดตามดูทำลาย" Then
            cboType.ListIndex = 3
            Label13.Caption = "/55"
        Else: cboType.ListIndex = 0
        Label13.Caption = ""
        End If
        If Val(.FullPoint) > 0 Then
            Label13.Caption = "/" & .FullPoint
        End If
        TxtRemark.Text = .Remark
        
        TxtPoint.Text = .Point
          lblAverage.Caption = Format(.Point, "###.#0") & "%"
'        If Label13.Caption <> "" Then
'            tPoint = Replace(Label13.Caption, "/", "")
'            Ave = (Trim(TxtPoint.Text) * 100) / tPoint
'            lblAverage.Caption = Format(Ave, "###.#0") & "%"
'        End If
        txtSale2.Text = .SaleStaff
        If .isComplete = True Then
            ChkIsComplete.Value = 1
        Else
            ChkIsComplete.Value = 0
        End If
        If .LookOnline = True Then
            ChkOnline.Value = 1
        Else
            ChkOnline.Value = 0
        End If
        If .IsInform = True Then
            ChkInform.Value = 1
        Else
            ChkInform.Value = 0
        End If
        If .IsNewCust = True Then
            ChkNewCust.Value = 1
        Else
            ChkNewCust.Value = 0
        End If
        If .isUse = True Then
            ChkUse.Value = 1
        Else
            ChkUse.Value = 0
        End If
        If .IsNoCome = True Then
            ChkNoCome.Value = 1
        Else
            ChkNoCome.Value = 0
        End If
        If .CheckList = True Then
            ChkCheckList.Value = 1
        Else
            ChkCheckList.Value = 0
        End If
        If .EmpPPE = True Then
            ChkPPE.Value = 1
        Else
            ChkPPE.Value = 0
        End If
        If .FieldClean = True Then
            ChkClean.Value = 1
        Else
            ChkClean.Value = 0
        End If
        If .CCTV = True Then
            ChkCCTV.Value = 1
        Else
            ChkCCTV.Value = 0
        End If
        If .EmpReady = True Then
            ChkReady.Value = 1
        Else
            ChkReady.Value = 0
        End If
        If .WorkProcess = True Then
            ChkProcess.Value = 1
        Else
            ChkProcess.Value = 0
        End If
        If .Welcome = True Then
            ChkWelcome.Value = 1
        Else
            ChkWelcome.Value = 0
        End If
        ShowCustStaff
        ShowRecStaff
    End With
    txtCustomer.Enabled = False
End Sub
Private Sub ShowCustStaff()
    Dim tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    With myCustVisit_
        tSql = "SELECT * FROM tbCustomerVisitDetail " & vbCrLf
        tSql = tSql & "where VisitID= '" & .id & "' " & vbCrLf
        tSql = tSql & "order by ID"
        Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
        If rs.RecordCount > 0 Then
            With grCustomer.getGridObj
                For i = 1 To rs.RecordCount
                    .TextMatrix(i, 0) = rs!id
                    .TextMatrix(i, 1) = rs!VisitorName
                    .TextMatrix(i, 2) = rs!Position
                rs.MoveNext
                Next
            End With
        End If
    End With
End Sub
Private Sub ShowRecStaff()
    Dim tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    With myCustVisit_
        tSql = "SELECT * FROM vw_Reception " & vbCrLf
        tSql = tSql & "where VisitID= '" & .id & "' " & vbCrLf
        tSql = tSql & "order by dept,StaffName"
        Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
        If rs.RecordCount > 0 Then
            With grStaff.getGridObj
                For i = 1 To rs.RecordCount
                    .TextMatrix(i, 0) = rs!id
                    .TextMatrix(i, 1) = rs!StaffName
                    .TextMatrix(i, 2) = rs!Dept
                    .TextMatrix(i, 3) = rs!StaffID
                rs.MoveNext
                Next
            End With
        End If
    End With
End Sub
Function CheckDigitOnly(ByVal Index As Integer) As Integer
'/ ---------------------------------------------------------------------------
        Select Case Index
            Case 48 To 57 ' ??? 0 - 9
            Case 8, 13 ' Backspace = 8, Enter = 13
            Case Else
                Index = 0
        End Select
        CheckDigitOnly = Index
    End Function
Private Sub TxtMail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub txtMail_Validate(Cancel As Boolean)
    myCustVisit_.Email = Trim(TxtMail.Text)
End Sub

Private Sub txtNameDisplay_Validate(Cancel As Boolean)
    myCustVisit_.NameDisplay = Trim(txtNameDisplay.Text)
End Sub

Private Sub TxtPoint_Change()
    If Not IsNumeric(TxtPoint.Text) Then SendKeys "{backspace}"
End Sub

Private Sub TxtPoint_Validate(Cancel As Boolean)
    Dim AveragePoint As Double
    Dim TotalPoint As Double
    
    myCustVisit_.Point = Val(TxtPoint.Text)
    TotalPoint = Val(Replace(Label13.Caption, "/", ""))
    
    If TotalPoint = 0 Then
        lblAverage.Caption = "N/A"
    Else
        AveragePoint = (Val(TxtPoint.Text) / TotalPoint) * 100
        lblAverage.Caption = Format(AveragePoint, "###.#0") & "%"
    End If
End Sub

Private Sub TxtPosition_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub TxtPosition_Validate(Cancel As Boolean)
    myCustVisit_.Position = Trim(TxtPosition.Text)
End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub txtRemark_Validate(Cancel As Boolean)
    myCustVisit_.Remark = Trim(TxtRemark.Text)
End Sub

Private Sub txtSale2_Change()
'    If Trim(txtSale2.Text) = "" Then Exit Sub
'    If curSaleStaff_ Is Nothing Then Set curSaleStaff_ = New BWGCompanyStaff
'    If Trim(txtSale2.Text) = "" And Trim(curSaleStaff_.StaffFName & " " & curSaleStaff_.StaffLName) <> Trim(txtSale2.Text) Then
'        Set curSaleStaff_ = Nothing
'        Set curSaleStaff_ = New BWGCompanyStaff
'        With MyCustomer_
'            .Sale2ID = ""
'            .Sale2Name = ""
'        End With
'    Else
'        If Trim(curSaleStaff_.StaffFName & " " & curSaleStaff_.StaffLName) <> Trim(txtSale2.Text) Then
'            Set curSaleStaff_ = Nothing
'            Set curSaleStaff_ = New BWGCompanyStaff
'            With MyCustomer_
'                .Sale2ID = ""
'                .Sale2Name = ""
'            End With
'            Load frmStaffSearch
'            With frmStaffSearch
'                .getCriteria (txtSale2.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If MyCustomer_.Sale2Name = "" Then txtSale2.Text = ""
'        End If
'    End If
End Sub
'Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
'    Set curSaleStaff_ = selStaff
'    With curSaleStaff_
'        txtSale2.Text = .StaffFName & " " & .StaffLName
'        If isAssigned Then
'            MyCustomer_.Sale2Name = .StaffFName & " " & .StaffLName
'            MyCustomer_.Sale2ID = .ID
'        End If
'    End With
'End Sub

Private Sub TxtTel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub txtTel_Validate(Cancel As Boolean)
    myCustVisit_.Tel = Trim(TxtTel.Text)
End Sub

Private Sub TxtTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub TxtTime_Validate(Cancel As Boolean)
    myCustVisit_.VisitTime = Trim(TxtTime.Text)
End Sub

Private Sub TxtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckDigitOnly(KeyAscii)
End Sub

Private Sub TxtTotal_Validate(Cancel As Boolean)
        myCustVisit_.TotalPeople = Val(TxtTotal.Text)
End Sub
