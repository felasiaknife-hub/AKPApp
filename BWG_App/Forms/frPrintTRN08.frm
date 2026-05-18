VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN08 
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   -60
      TabIndex        =   8
      Top             =   780
      Width           =   6135
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   4455
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   1260
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Caption         =   "รายงานสรุปใบเมนิเฟสที่ไม่มีการขนเวสต์ (TRN08)"
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
      TabIndex        =   7
      Top             =   60
      Width           =   4320
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintTRN08.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN08"
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
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    PopulateCompany cboCompany, cboIndex
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, subStr As String
    With rpt1
        .Reset
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\TRN08.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_MenifestDetailSearchCheckNotWaste V1" & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate>='" & dtFrom.ValueYMD & "' AND WorkDate<='" & dtTo.ValueYMD & "'" & vbCrLf
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode) = "AKP" Then
                    tmpSql = tmpSql & "AND (CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID) & "' "
                    'tmpSql = tmpSql & "OR WasteType IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI'))" & vbCrLf
                    tmpSql = tmpSql & "AND DocNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode) = "PST" Then
                    tmpSql = tmpSql & "AND (CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID) & "' "
                Else
                    tmpSql = tmpSql & "AND CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID) & "' "
                    tmpSql = tmpSql & "AND DocNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode) = "AKP" Then
                    tmpSql = tmpSql & "AND (CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID) & "' "
                    'tmpSql = tmpSql & "OR WasteType IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI'))" & vbCrLf
                    tmpSql = tmpSql & "AND DocNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode) = "PST" Then
                    tmpSql = tmpSql & "AND (CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID) & "' "
                Else
                    tmpSql = tmpSql & "AND CompanyID='" & Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID) & "' "
                    tmpSql = tmpSql & "AND DocNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        End If
        
        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & "AND UnderSaleStaffID='" & curStaff_.ID & "' "
        End If
        
'        tmpSql = tmpSql & " AND (FailCheck=1)" & vbCrLf
        tmpSql = tmpSql & "AND (SELECT Count(T1.MenifestID) From tbMenifestDetail T1 WHERE T1.MenifestID = V1.MenifestID) = (SELECT Count(T2.MenifestID) From tbJobFail T2 WHERE T2.MenifestID = V1.MenifestID AND ISNULL(FailCheck,0) = 1)" & vbCrLf
        tmpSql = tmpSql & "ORDER BY WorkDate, DocNo"
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
'        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
'        .Formulas(3) = "CarTypeName=" & Chr(34) & cboCarType.Text & Chr(34)
        .WindowTitle = "รายงานสรุปสถานะใบเมนิเฟส"
        .WindowState = crptMaximized
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
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

