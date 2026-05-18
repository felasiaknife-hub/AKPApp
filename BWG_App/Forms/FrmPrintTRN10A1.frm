VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPrintTRN10A1 
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   6105
      Begin VB.TextBox TxtCarReg 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1530
         Width           =   4425
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   4425
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1140
         Width           =   4425
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ทะเบียนรถ"
         Height          =   195
         Index           =   5
         Left            =   450
         TabIndex        =   13
         Top             =   1575
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ขนโดย"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   11
         Top             =   810
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   9
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4155
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2775
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2775
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3195
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2775
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   5910
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานระยะทางการขนกากของรถแต่ละคัน (TRN10A-1)"
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
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   4890
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "FrmPrintTRN10A1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "FrmPrintTRN10A1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim coltruckType_ As Collection
Dim curStaff_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Public txt As String
Dim curTrans2_ As BWGTransportVendor
Dim curTruckSubType_ As BWGTruckSubType

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub cmdPreview_Click()
    If cboCarType.ListIndex = -1 Then
        MsgBox "กรุณาระบุประเภทรถ", vbExclamation
        cboCarType.SetFocus
        Exit Sub
    End If
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    If cboCarType.ListIndex = -1 Then
        MsgBox "กรุณาระบุประเภทรถ", vbExclamation
        cboCarType.SetFocus
        Exit Sub
    End If
    PrintMain 1
End Sub

Private Sub Form_Load()
    PopulateCarType
    PopulateCompany cboCompany, cboIndex
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = 1
End Sub

Private Sub PopulateCarType()
Dim x%, tmpS As New BWGSearchManager
    cboCarType.Clear
    Set coltruckType_ = tmpS.TruckTypeSearch
    cboCarType.AddItem "==เลือกทุกประเภทรถ=="
    For x = 1 To coltruckType_.Count
        cboCarType.AddItem coltruckType_(x).TypeDesc
    Next
    If cboCarType.ListCount > 0 Then
        cboCarType.ListIndex = 0
    End If
End Sub
Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    Dim tmpRec As New ADODB.Recordset
Dim tmpSql1 As String

    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\TRN10A1.rpt"
        tmpSql = ""
        tmpSql = "select subtypedesc,CarRegisID,DriverName,sum(LengthNum) as LengthNum,sum(RunDay) as RunDay,sum(RunHour) as RunHour,sum(CarRate) as CarRate" & vbCrLf
        tmpSql = tmpSql & " from vw_MenifestSearch " & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate >= '" & dtFrom.ValueYMD & "' AND WorkDate <= '" & dtTo.ValueYMD & "'" & vbCrLf
        If cboCarType.ListIndex > 0 Then
            tmpSql = tmpSql & "AND TruckTypeID = '" & coltruckType_(cboCarType.ListIndex).ID & "'" & vbCrLf
        End If
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        End If
        If TxtCarReg.Text <> "" Then
                tmpSql = tmpSql & "AND CarRegisID like '%" & TxtCarReg.Text & "%'" & vbCrLf
        End If
        tmpSql = tmpSql & "group BY subtypedesc,CarRegisID,DriverName"
        Set tmpRec = GetRS(tmpSql, adOpenForwardOnly, adLockReadOnly)
            If tmpRec.RecordCount = 0 Then Screen.MousePointer = 0: Exit Sub
        DBConnExc "Delete TbTmpLenght "
    Do While Not tmpRec.EOF
        Dim LN, RD, RH, CR
        If IsNull(tmpRec!LengthNum) = True Then
            LN = "0"
        Else
            LN = tmpRec!LengthNum
        End If
        If IsNull(tmpRec!RunDay) = True Then
            RD = "0"
        Else
            RD = tmpRec!RunDay
        End If
        If IsNull(tmpRec!RunHour) = True Then
            RH = "0"
        Else
            RH = tmpRec!RunHour
        End If
        If IsNull(tmpRec!CarRate) = True Then
            CR = "0"
        Else
            CR = tmpRec!CarRate
        End If
        
   DBConnExc "Insert into TbTmpLenght (subtypedesc,CarRegisID,DriverName,LengthNum,RunDay,RunHour,CarRate) " & _
                            " Values ('" & tmpRec!SubTypeDesc & "','" & tmpRec!CarRegisID & "','" & tmpRec!DriverName & "'," & LN & "," & RD & "" & _
                            "," & RH & "," & CR & ")"
        tmpRec.MoveNext
    Loop
        
        
        tmpSql1 = ""
        tmpSql1 = "select * " & vbCrLf
        tmpSql1 = tmpSql1 & " from TbTmpLenght " & vbCrLf

        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .WindowTitle = "รายงานระยะทางการขนกาก (TRN10A-1)"
        .SQLQuery = tmpSql1
        .Destination = printType
        .PrintReport
'        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub


