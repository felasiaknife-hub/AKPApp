VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPrintTRN01 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   30
      TabIndex        =   7
      Top             =   780
      Width           =   5985
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   1530
         Width           =   4425
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   750
         Width           =   4425
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   4425
      End
      Begin Project1.ctlDate dtFrom 
         Height          =   345
         Left            =   1140
         TabIndex        =   0
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BackColor       =   -2147483633
      End
      Begin Project1.ctlDate dtTo 
         Height          =   345
         Left            =   3810
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   14
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ขนโดย"
         Height          =   195
         Index           =   4
         Left            =   540
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
         Left            =   300
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
         Left            =   270
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
         Left            =   3180
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
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
      Caption         =   "รายงานการขนกากแยกตามประเภทรถ"
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
      TabIndex        =   6
      Top             =   60
      Width           =   3210
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintTRN01.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frPrintTRN01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim coltruckType_ As Collection
Dim curStaff_ As BWGCompanyStaff

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
    PopulateCarType
    PopulateCompany
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Private Sub PopulateCarType()
Dim x%, tmpS As New BWGSearchManager
    cboCarType.Clear
    Set coltruckType_ = tmpS.TruckTypeSearch
    For x = 1 To coltruckType_.Count
        cboCarType.AddItem coltruckType_(x).TypeDesc
    Next
    If cboCarType.ListCount > 0 Then
        cboCarType.ListIndex = 0
    End If
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\reports\TRN01.rpt"
        tmpSql = ""
        tmpSql = "Select vw_RptTRN01.*,vw_RptTRN01Det.* from vw_RptTRN01 inner join vw_RptTRN01Det on vw_RptTRN01.JobID=vw_RptTRN01Det.JobID " & vbCrLf
        tmpSql = tmpSql & "Where WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And TruckTypeID='" & coltruckType_(cboCarType.ListIndex + 1).ID & "' And CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID) & "' "
        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "' "
        End If
        tmpSql = tmpSql & vbCrLf & " Order By CompanyName,TruckTypeDesc"
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(3) = "CarTypeName=" & Chr(34) & cboCarType.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
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
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtSale.Text)
                .setCallerForm Me
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

