VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCRBigbag 
   Caption         =   "รายงาน Bigbag"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายงานการใช้ Bigbag"
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   15
      TabIndex        =   3
      Top             =   735
      Width           =   5925
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1110
         Width           =   4395
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   735
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   11
         Top             =   1170
         Width           =   375
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
         TabIndex        =   10
         Top             =   795
         Width           =   735
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   8
         Top             =   420
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2370
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2370
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2370
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
      Caption         =   "รายงานการใช้ Bigbag"
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
      Width           =   1845
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "FrmCRBigbag.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "FrmCRBigbag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colZone_ As Collection
Dim coltruckType_ As Collection
Dim curStaff_ As BWGCompanyStaff
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
Dim tmpWaste As BWGWasteDataCR
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
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
Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If Label1(1).Caption = "รายงานการใช้ Bigbag" Then
        .ReportFileName = App.Path & "\reports\CRBigbag.rpt"
        Else
        .ReportFileName = App.Path & "\reports\CRCovidBin.rpt"
        End If
        tmpSql = ""
        If Label1(1).Caption = "รายงานการใช้ Bigbag" Then
        tmpSql = "SELECT * FROM vw_RptBigbag " & vbCrLf
        Else
        tmpSql = "SELECT * FROM vw_RptCovidBin " & vbCrLf
        End If
        tmpSql = tmpSql & "WHERE WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' "

        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & " AND UnderSaleStaffID='" & curStaff_.ID & "'"
        End If
        tmpSql = tmpSql & vbCrLf & "ORDER BY WorkDate,CompanyName"
        .Formulas(0) = "FromDate=" & Chr(34) & "ข้อมูลวันที่ " & " " & dtFrom.ValueDMY & " " & " ถึงวันที่ " & " " & dtTo.ValueDMY & Chr(34)
'        .Formulas(0) = "PrintedDate=" & Chr(34) & ข้อมูลวันที่ & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

