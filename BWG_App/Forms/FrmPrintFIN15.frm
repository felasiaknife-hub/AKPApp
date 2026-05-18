VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmPrintFIN15 
   Caption         =   "รายได้ค่ากำจัด ค่าถัง ค่าถุง ค่าคนงาน"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkprintEX 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   255
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
      Begin VB.CheckBox ChkDateCust 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   1920
         Width           =   135
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   4395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   1500
         Width           =   4365
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   60
         Top             =   480
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
         Caption         =   "เฉพาะลูกค้าใหม่ดิว 90 วัน"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท :"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   8
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า :"
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   7
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   6
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Caption         =   "เลือก ในรูปแบบที่สามารถ Export Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "รายงานค่ากำจัด ค่าถัง ค่าถุง ค่าคนงาน FIN15"
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
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   3315
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "FrmPrintFIN15.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "FrmPrintFIN15"
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
'    CanAccess = False: CanUpdate = False
'    SetPermission Me, CanAccess, CanUpdate
'    If CanAccess = False Then
'        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
'        Unload Me
'        Exit Sub
'    End If
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    PopulateCompany cboCompany, cboIndex
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
    If chkprintEX.Value = 1 Then
       .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\FIN15-1.rpt"
    Else
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\FIN15.rpt"
    End If

        tmpSql = ""
        tmpSql = "SELECT Distinct * FROM vw_RptFin15 " & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate >='" & dtFrom.ValueYMD & "' And WorkDate <='" & dtTo.ValueYMD & "'"
        
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
            End If
        End If
        If ChkDateCust.Value = 1 Then
            tmpSql = tmpSql & " And SetDay = '90' "
        End If
        
        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
         tmpSql = tmpSql & vbCrLf & " ORDER BY WorkDate,TimeTableNo"
        .Formulas(0) = "ShowDateRange=" & Chr(34) & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
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





