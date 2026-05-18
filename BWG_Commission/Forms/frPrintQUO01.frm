VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPrintQUO01 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   60
      TabIndex        =   8
      Top             =   780
      Width           =   5925
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1170
         Width           =   4395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   750
         Width           =   4365
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
         Left            =   3750
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทที่ออก"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   12
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จนท. ผู้ออก"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   810
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ออก"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   9
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2490
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2490
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2490
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
      Caption         =   "รายงานเปรียบเทียบจำนวนใบเสนอราคา (QUO01)"
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
      TabIndex        =   7
      Top             =   60
      Width           =   4275
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintQUO01.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frPrintQUO01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    PopulateCompany
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\reports\QUO01.rpt"
        tmpSql = ""
        tmpSql = "Select * from vw_RptQUO01 " & vbCrLf
        tmpSql = tmpSql & "Where IssueDate>='" & dtFrom.ValueYMD & "' And IssueDate<='" & dtTo.ValueYMD & "' "
        If Not curStaff_ Is Nothing Then
            If Trim(curStaff_.ID) <> "" Then
                tmpSql = tmpSql & " And StaffID='" & curStaff_.ID & "' "
            End If
        End If
        If cboCompany.ListIndex > 0 Then
            tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex).ID & "' "
        End If
        tmpSql = tmpSql & vbCrLf & " Order By ZoneName,SDName,StaffFName,StaffLName"
        .Formulas(0) = "DateRange=" & Chr(34) & "ช่วงวันที่ออกใบเสนอราคา " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .Formulas(2) = "ShowWorkCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtCustomer.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtCustomer.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtCustomer.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtCustomer.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtCustomer.Text = ""
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
        txtCustomer.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    cboCompany.AddItem "ทุกบริษัท"
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

