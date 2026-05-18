VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPrintORD06 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   60
      TabIndex        =   8
      Top             =   810
      Width           =   5925
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   4455
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   1110
         Width           =   4425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปี"
         Height          =   195
         Index           =   0
         Left            =   3570
         TabIndex        =   11
         Top             =   420
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   10
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เดือน"
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   9
         Top             =   420
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3060
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
      Caption         =   "รายงานยอดขายตามผู้แทนขาย"
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
      Width           =   2610
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintORD06.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frPrintORD06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_ As String

Private Sub cmdClose_Click()
    Dim Frm As Form
    For Each Frm In Forms
        Unload Frm
    Next
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
Dim x%
    sessionID_ = GetGUID
    PopulateCompany
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboMonth
        .AddItem "มกราคม"
        .AddItem "กุมภาพันธ์"
        .AddItem "มีนาคม"
        .AddItem "เมษายน"
        .AddItem "พฤษภาคม"
        .AddItem "มิถุนายน"
        .AddItem "กรกฎาคม"
        .AddItem "สิงหาคม"
        .AddItem "กันยายน"
        .AddItem "ตุลาคม"
        .AddItem "พฤศจิกายน"
        .AddItem "ธันวาคม"
        .ListIndex = CInt(Mid(TodayDate, 6, 2)) - 1
    End With
    With cboYear
        For x = 1 To 10
            .AddItem (CInt(Left(TodayDate, 4)) - 10 + x)
        Next
        SetTextToCombo Left(TodayDate, 4), cboYear
    End With
End Sub

Private Sub PrintMain(printType As Integer)
On Error GoTo ErrD
Dim tmpSql As String, iCount As Long
Dim tmpCon As ADODB.Connection, tmpRec As New ADODB.Recordset
    Set tmpCon = DBManager.GetDBConnection
    tmpSql = ""
    tmpSql = "Select * from Com_vwCommission Where ForYear = '" & cboYear.Text & "' AND ForMonth = '" & Format(cboMonth.ListIndex + 1, "0") & "'"
    tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'"
    If Trim(txtSale.Text) <> "" Then
        tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "'"
    End If
    tmpSql = tmpSql & " AND ISNull(SaleStaffID,'') <> ''"
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\Reports\ORD06.rpt"
        
        .Formulas(0) = "WorkDate='ประจำเดือน  " & cboMonth.Text & "  " & cboYear.Text & "'"
        .Formulas(1) = "Company='" & cboCompany.Text & "'"
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    Exit Sub
ErrD:
    MsgBox Err.Description
    Resume Next
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBManager.GetDBConnection.Execute "Delete tbPrintOrd01 Where SessionID='" & sessionID_ & "'"
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

