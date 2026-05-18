VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   30
      TabIndex        =   10
      Top             =   780
      Width           =   5985
      Begin VB.ComboBox cboShowRecords 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   4425
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1890
         Width           =   4425
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   4425
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   4425
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "************"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   16
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   15
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ขนโดย"
         Height          =   195
         Index           =   4
         Left            =   540
         TabIndex        =   14
         Top             =   810
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   12
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   11
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   180
      Top             =   3420
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
      Caption         =   "รายงานการขนกากที่ขาดทุนค่าขนส่ง (TRN14)"
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
      TabIndex        =   9
      Top             =   60
      Width           =   3960
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintTRN14.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Public isMD As Boolean

Private Sub cmdClose_Click()
    Unload Me
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
    PopulateCarType cboCarType
    PopulateCompany cboCompany, cboIndex
    With cboShowRecords
        .AddItem "==แสดงทุกรายการ=="
        .AddItem "แสดงเฉพาะรายการที่ค่าขนส่งแยก น้อยกว่า ค่าขนส่งที่จ่ายเวนเดอร์"
        .AddItem "แสดงเฉพาะรายการที่ค่าขนส่งแฝง = 0"
        .ListIndex = 0
    End With
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = 1
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If isMD = False Then
            .ReportFileName = App.Path & "\reports\TRN13.rpt"
        Else
            .ReportFileName = App.Path & "\reports\TRN13_MD.rpt"
        End If
        tmpSql = "SELECT * FROM vw_MenifestSearch" & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'" & vbCrLf
        tmpSql = tmpSql & "AND isCanceled <> 'Y' AND isClosed = 'Y'" & vbCrLf
'        tmpSql = tmpSql & "AND ActualQty > 0" & vbCrLf

'        tmpSql = tmpSql & "AND CustomerName LIKE '%นวนคร%')" & vbCrLf
'        tmpSql = tmpSql & "OR (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'" & vbCrLf
'        tmpSql = tmpSql & "AND isCanceled <> 'Y' AND isClosed = 'Y'" & vbCrLf
'        tmpSql = tmpSql & "AND SiteAumphurName = 'คลองหลวง')" & vbCrLf
        
        If cboCarType.ListIndex > 0 Then tmpSql = tmpSql & "AND TruckTypeID='" & coltruckType_(cboCarType.ListIndex).ID & "'" & vbCrLf
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        End If
        
        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "' " & vbCrLf
        End If
        tmpSql = tmpSql & "ORDER BY CompanyNameTH, TruckTypeDesc, CustomerName, TimeTableNo, IncollectFee Desc"
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(3) = "ReportTitle = " & Chr(34) & cboCarType.Text & Chr(34)
        If cboCarType.ListIndex <= 0 Then .Formulas(4) = "ShowGroup1 = " & Chr(34) & "False" & Chr(34)
        If cboShowRecords.ListIndex > 0 Then .Formulas(5) = "ShowRecords = " & Chr(34) & cboShowRecords.ListIndex & Chr(34)
        .WindowTitle = "รายงานการขนกากแยกตามประเภทรถ (TRN13)"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

