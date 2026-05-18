VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintFIN09P2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F086"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      TabIndex        =   15
      Top             =   780
      Width           =   6915
      Begin VB.ComboBox cboApprove 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   4395
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ค้นหาตามเวลาปิดงาน"
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   2340
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   4680
         TabIndex        =   9
         Top             =   2355
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM"
         Format          =   85000194
         CurrentDate     =   40821
      End
      Begin VB.ComboBox cboClosedType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1500
         Width           =   4395
      End
      Begin VB.ComboBox cboDateType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1755
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   1755
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   4680
         TabIndex        =   10
         Top             =   2715
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM"
         Format          =   85000194
         CurrentDate     =   40821
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อนุมัติการปิดงาน"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   22
         Top             =   1980
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงเวลาปิดงาน"
         Height          =   195
         Index           =   7
         Left            =   3540
         TabIndex        =   21
         Top             =   2355
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทผู้ปิดงาน"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานะ"
         Height          =   195
         Index           =   5
         Left            =   3480
         TabIndex        =   19
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทกาก"
         Height          =   195
         Index           =   4
         Left            =   540
         TabIndex        =   18
         Top             =   1140
         Width           =   855
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   16
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4080
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทกาก"
      Height          =   195
      Index           =   8
      Left            =   540
      TabIndex        =   23
      Top             =   2400
      Width           =   855
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
      Caption         =   "รายงานการขนแยกตาม Mainfest และการตรวจสอบเอกสาร (FIN09-2)"
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
      TabIndex        =   14
      Top             =   60
      Width           =   5940
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintFIN09_2.frx":0000
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintFIN09P2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    DTPicker1.Visible = CBool(Check1.Value)
    DTPicker2.Visible = CBool(Check1.Value)
    Label1(7).Visible = CBool(Check1.Value)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x As Integer
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
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
    With CboType
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

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
Dim selDate As String, DateDesc As String, Date1 As String, Date2 As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\FIN09_2.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_RptFIN9_2 AS vw_RptFIN09 " & vbCrLf
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
        
        If CboType.ListIndex > 0 Then
            tmpSql = tmpSql & " And WasteType='" & CboType.Text & "' "
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

        If Check1.Value = 1 Then tmpSql = tmpSql & " AND CloseTime BETWEEN '" & Format(DTPicker1.Value, "HH:MM") & "' AND '" & Format(DTPicker2.Value, "HH:MM") & "'"

        tmpSql = tmpSql & "ORDER BY PriorityLevel, DocNo"
        
        .Formulas(0) = "DateRange = " & Chr(34) & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & " [เวลาปิดงาน " & Format(DTPicker1.Value, "HH:MM") & " - " & Format(DTPicker2.Value, "HH:MM") & "] " & Chr(34)
        .Formulas(1) = "PrintedDate = " & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .Formulas(2) = "ShowCompany = " & Chr(34) & cboCompany.Text & Chr(34)
        If CboType.ListIndex = 0 Then
            .Formulas(3) = "ShowWasteType = " & Chr(34) & "ทุกประเภทของเสีย" & Chr(34)
        Else
            .Formulas(3) = "ShowWasteType = " & Chr(34) & CboType.Text & Chr(34)
        End If
        .Formulas(4) = "ShowDateType = " & Chr(34) & DateDesc & Chr(34)
        .Formulas(5) = "ShowApprove = " & Chr(34) & "สถานะการอนุมัติ = " & cboApprove.Text & Chr(34)
        .WindowShowCloseBtn = True
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

