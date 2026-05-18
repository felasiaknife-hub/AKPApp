VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmFin09DIW 
   Caption         =   "กำหนดข้อมูลที่ให้กรมฯ ดูได้"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4095
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4110
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "บันทึก"
      Height          =   795
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4110
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   810
      Width           =   6915
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   4395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1755
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1755
      End
      Begin VB.ComboBox cboDateType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox cboClosedType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   4395
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ค้นหาตามเวลาปิดงาน"
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   2340
         Width           =   1935
      End
      Begin VB.ComboBox cboApprove 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1920
         Width           =   4395
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   2355
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM"
         Format          =   62914562
         CurrentDate     =   40821
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   4680
         TabIndex        =   11
         Top             =   2715
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM"
         Format          =   62914562
         CurrentDate     =   40821
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   18
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทที่ขน"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทกาก"
         Height          =   195
         Index           =   4
         Left            =   540
         TabIndex        =   16
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานะ"
         Height          =   195
         Index           =   5
         Left            =   3480
         TabIndex        =   15
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทผู้ปิดงาน"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงเวลาปิดงาน"
         Height          =   195
         Index           =   7
         Left            =   3540
         TabIndex        =   13
         Top             =   2355
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อนุมัติการปิดงาน"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   12
         Top             =   1980
         Width           =   1170
      End
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กำหนดข้อมูลสำหรับรายงานที่กรมฯ"
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
      TabIndex        =   23
      Top             =   90
      Width           =   3075
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   5910
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทกาก"
      Height          =   195
      Index           =   8
      Left            =   540
      TabIndex        =   22
      Top             =   2430
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "FrmFin09DIW.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "FrmFin09DIW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fin092 As New AKPFin09DIW
Private Sub Check1_Click()
    DTPicker1.Visible = CBool(Check1.Value)
    DTPicker2.Visible = CBool(Check1.Value)
    Label1(7).Visible = CBool(Check1.Value)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    SaveData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x As Integer
    PopulateCompany cboCompany, cboIndex
    With cboDateType
        .AddItem "วันที่ขน"
        .AddItem "วันที่ปิด Job"
        .AddItem "วันที่อนุมัติ"
        .ListIndex = 0
    End With
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboType
        .AddItem "ทุกประเภท"
        For x = 1 To WasteTypes.Count
            .AddItem WasteTypes(x).WasteType
        Next
        .ListIndex = 0
    End With
    With cboStatus
        .AddItem "ทุกประเภท"
        .AddItem "Active"
        .AddItem "Canceled"
        .ListIndex = 0
    End With
    With cboClosedType
        .AddItem "เลือกทั้งหมด"
        .AddItem "ปิดงานโดยเจ้าหน้าที่ขนส่ง"
        .AddItem "ปิดงานโดย Application"
        .AddItem "ปิดงานโดย System Database"
        .ListIndex = 0
    End With

    With cboApprove
        .AddItem "เลือกทั้งหมด"
        .AddItem "อนุมัติการปิดงาน"
        .AddItem "ยังไม่อนุมัติการปิดงาน"
        .ListIndex = 0
    End With
    
    DTPicker1.Hour = "00"
    DTPicker1.Minute = "00"
    DTPicker2.Hour = "23"
    DTPicker2.Minute = "59"
    
    DTPicker1.Visible = CBool(Check1.Value)
    DTPicker2.Visible = CBool(Check1.Value)
    Label1(7).Visible = CBool(Check1.Value)
End Sub
Private Sub SaveData()
Dim tmpSql, Sql As String
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim i, x As Integer
Dim selDate As String, DateDesc As String, Date1 As String, Date2 As String
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_RptFIN9_2 " & vbCrLf
        Date1 = dtFrom.ValueYMD
        Date2 = dtTo.ValueYMD
        If cboDateType.ListIndex = 0 Then
            selDate = "WorkDate"
            DateDesc = "ช่วงวันที่ขนกาก"
        ElseIf cboDateType.ListIndex = 1 Then
            selDate = "CloseDate"
            DateDesc = "ช่วงวันที่ปิดงาน"
        ElseIf cboDateType.ListIndex = 2 Then
            selDate = "ApproveDate"
            DateDesc = "ช่วงวันที่อนุมัติ"
            Date1 = dtFrom.YearValue - 543 & "/" & Format(dtFrom.MonthValue, "00") & "/" & Format(dtFrom.DayValue, "00")
            Date2 = dtTo.YearValue - 543 & "/" & Format(dtTo.MonthValue, "00") & "/" & Format(dtTo.DayValue, "00")
        End If
        tmpSql = tmpSql & "WHERE " & selDate & ">='" & Date1 & "' And " & selDate & "<='" & Date2 & "' "
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE '%PST%'"
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                Else
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE '%PST%'"
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE '%PST%'"
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                Else
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE '%PST%'"
                End If
            End If
        End If
        If cboType.ListIndex > 0 Then
            tmpSql = tmpSql & " And WasteType='" & cboType.Text & "' "
        End If
        If cboStatus.ListIndex = 1 Then
            tmpSql = tmpSql & " And (isCanceled<>'Y' Or isCanceled is null)"
        ElseIf cboStatus.ListIndex = 2 Then
            tmpSql = tmpSql & " And (isCanceled='Y')"
        End If
        Select Case cboClosedType.ListIndex
            Case 1
                tmpSql = tmpSql & " AND CloseStaffID NOT IN (SELECT StaffID FROM tbCompanyStaff WHERE StaffCode = 'SYS')" & vbCrLf
            Case 2
                tmpSql = tmpSql & " AND CloseStaffID IN (SELECT StaffID FROM tbCompanyStaff WHERE StaffFName = 'Application')" & vbCrLf
            Case 3
                tmpSql = tmpSql & " AND CloseStaffID IN (SELECT StaffID FROM tbCompanyStaff WHERE StaffFName = 'System')" & vbCrLf
        End Select
        Select Case cboApprove.ListIndex
            Case 1
                tmpSql = tmpSql & " AND isApproved = 'Y'" & vbCrLf
            Case 2
                tmpSql = tmpSql & " AND isApproved = 'N'" & vbCrLf
        End Select
        tmpSql = tmpSql & " AND nonorhz <> 'DOC' " & vbCrLf
        If Check1.Value = 1 Then tmpSql = tmpSql & " AND CloseTime BETWEEN '" & Format(DTPicker1.Value, "HH:MM") & "' AND '" & Format(DTPicker2.Value, "HH:MM") & "'"
        tmpSql = tmpSql & "ORDER BY PriorityLevel, DocNo"
        Set rs = GetRS(tmpSql, adOpenForwardOnly, adLockOptimistic)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For i = 1 To rs.RecordCount
                With Fin092
                    .docNo = rs!docNo
                    If IsNull(rs!MenifestNo) = False Then .MenifestNo = rs!MenifestNo
                    If IsNull(rs!WasteGenNo) = False Then .WasteGenNo = rs!WasteGenNo
                    If IsNull(rs!TransRegisNo1) = False Then .TransRegisNo1 = rs!TransRegisNo1
                    If IsNull(rs!TransportName2) = False Then .TransportName2 = rs!TransportName2
                    If IsNull(rs!CarRegisID) = False Then .CarRegisID = rs!CarRegisID
                    If IsNull(rs!WorkDate) = False Then .WorkDate = rs!WorkDate
                    If IsNull(rs!IsCanceled) = False Then .IsCanceled = rs!IsCanceled
                    If IsNull(rs!isclosed) = False Then .isclosed = rs!isclosed
                    If IsNull(rs!DriverName) = False Then .DriverName = rs!DriverName
                    If IsNull(rs!JobDataCarID) = False Then .JobDataCarID = rs!JobDataCarID
                    If IsNull(rs!CompanyName) = False Then .CompanyName = rs!CompanyName
                    If IsNull(rs!CustAddr) = False Then .CustAddr = rs!CustAddr
                    If IsNull(rs!CustomerCode) = False Then .CustomerCode = rs!CustomerCode
                    If IsNull(rs!WasteName) = False Then .WasteName = rs!WasteName
                    If IsNull(rs!CustWeight) = False Then .CustWeight = rs!CustWeight
                    If IsNull(rs!DisposerWeight) = False Then .DisposerWeight = rs!DisposerWeight
                    If IsNull(rs!SubTypeDesc) = False Then .SubTypeDesc = rs!SubTypeDesc
                    If IsNull(rs!TruckTypeDesc) = False Then .TruckTypeDesc = rs!TruckTypeDesc
                    If IsNull(rs!FactoryRegisNo) = False Then .FactoryRegisNo = rs!FactoryRegisNo
                    If IsNull(rs!TreatmentRate) = False Then .TreatmentRate = rs!TreatmentRate
                    If IsNull(rs!EUCode) = False Then .EUCode = rs!EUCode
                    If IsNull(rs!PriorityLevel) = False Then .PriorityLevel = rs!PriorityLevel
                    If IsNull(rs!TreatmentUnitID) = False Then .TreatmentUnitID = rs!TreatmentUnitID
                    If IsNull(rs!RefME) = False Then .RefME = rs!RefME
                    If IsNull(rs!HaType) = False Then .HaType = rs!HaType
                    If IsNull(rs!ChargeWType) = False Then .ChargeWType = rs!ChargeWType
                    .Save
                End With
                rs.MoveNext
            Next
            MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว", vbOKOnly + vbInformation, "บันทึกข้อมูลสำเร็จ"
        Else
            MsgBox "ไม่พบข้อมูลที่ต้องการ", vbOKOnly + vbExclamation, "บันทึกข้อมูลไม่สำเร็จ"
        End If
End Sub


