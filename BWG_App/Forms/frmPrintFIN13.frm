VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintFIN13 
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   30
      TabIndex        =   3
      Top             =   870
      Width           =   6015
      Begin VB.OptionButton OptChek1 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1140
         TabIndex        =   18
         Top             =   780
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton OpChek2 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         Top             =   780
         Width           =   255
      End
      Begin VB.ComboBox cboCheck 
         Height          =   315
         ItemData        =   "frmPrintFIN13.frx":0000
         Left            =   1140
         List            =   "frmPrintFIN13.frx":0010
         TabIndex        =   15
         Top             =   2340
         Width           =   2355
      End
      Begin VB.ComboBox cboRE_Men 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   1920
         Width           =   2355
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1500
         Width           =   4365
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1140
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   7
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   60
         Top             =   120
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
         Caption         =   "วันที่ออกบิล"
         Height          =   195
         Index           =   6
         Left            =   2880
         TabIndex        =   21
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่วางบิล"
         Height          =   195
         Index           =   5
         Left            =   1380
         TabIndex        =   20
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "รูปแบบวางบิล"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "เจ้าหน้าที่"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   11
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   10
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   9
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วางบิลโดย"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   1200
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3780
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3780
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3780
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   1500
      TabIndex        =   19
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "รายงานรายละเอียดการวางบิล FIN13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   540
      Width           =   2715
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   -120
      Picture         =   "frmPrintFIN13.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintFIN13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer

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
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
        Exit Sub
    End If
    
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    PopulateCompany cboCompany, cboIndex
    PopulateStaffInfo cboRE_Men
    cboCheck.ListIndex = 0
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\FIN13.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_RptFIN01 " & vbCrLf
        
'        tmpSql = tmpSql & "WHERE IssuedDate>='" & dtFrom.ValueYMD & "' And IssuedDate<='" & dtTo.ValueYMD & "'"
         If OptChek1.Value = True Then
            tmpSql = tmpSql & "WHERE PayDueDate>='" & dtFrom.ValueYMD & "' And PayDueDate<='" & dtTo.ValueYMD & "'"
        End If
        If OpChek2.Value = True Then
            tmpSql = tmpSql & "WHERE issuedDate >='" & dtFrom.ValueYMD & "' And issuedDate <='" & dtTo.ValueYMD & "'"
        End If
        
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
            End If
        End If
        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
         If cboRE_Men.ListIndex > 0 Then
            tmpSql = tmpSql & " AND InfostaffNameRE = '" & cboRE_Men.Text & "' "
         End If
          If cboCheck.ListIndex = 1 Then
            tmpSql = tmpSql & " AND Inv_State = 'S' "
         End If
         If cboCheck.ListIndex = 2 Then
            tmpSql = tmpSql & " AND Inv_State = 'M' "
         End If
          If cboCheck.ListIndex = 3 Then
            tmpSql = tmpSql & " AND Inv_State = 'O "
         End If
        tmpSql = tmpSql & vbCrLf & " ORDER BY BillingNo"
        .Formulas(0) = "ShowDateRange=" & Chr(34) & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub




