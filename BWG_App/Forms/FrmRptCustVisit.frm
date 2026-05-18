VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRptCustVisit 
   Caption         =   "รายงานลูกค้าเข้าศูนย์"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   675
      Width           =   6135
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1125
         TabIndex        =   17
         Top             =   1755
         Width           =   4425
      End
      Begin VB.ComboBox cboCall 
         Height          =   315
         Left            =   3855
         TabIndex        =   15
         Top             =   1395
         Width           =   1725
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "FrmRptCustVisit.frx":0000
         Left            =   1140
         List            =   "FrmRptCustVisit.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   735
         Width           =   4455
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "FrmRptCustVisit.frx":0004
         Left            =   1140
         List            =   "FrmRptCustVisit.frx":0006
         TabIndex        =   6
         Top             =   1395
         Width           =   1500
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1065
         Width           =   4425
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   18
         Top             =   1815
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "การแจ้งล่วงหน้า"
         Height          =   195
         Index           =   4
         Left            =   2745
         TabIndex        =   16
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   14
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   465
         TabIndex        =   13
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วัตถุประสงค์"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   12
         Top             =   795
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   645
         TabIndex        =   10
         Top             =   1065
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2895
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2895
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2895
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   60
      Top             =   0
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
      Caption         =   "รายงานลูกค้าเข้าศูนย์ (PR01)"
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
      Left            =   180
      TabIndex        =   3
      Top             =   60
      Width           =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   270
      X2              =   5970
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   180
      X2              =   5820
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   -30
      Picture         =   "FrmRptCustVisit.frx":0008
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "FrmRptCustVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Private Sub Form_Load()
    With cboCustType
        .AddItem "== ทั้งหมด =="
        .AddItem "C,S"
        .AddItem "M,MS"
        .AddItem "B,E"
        .ListIndex = 0
    End With

    With cboCall
        .AddItem "==ทั้งหมด=="
        .AddItem "ที่แจ้งล่วงหน้า"
        .AddItem "ไม่ได้แจ้งล่วงหน้า"
        .ListIndex = 0
    End With

    With cboStatus
        .AddItem "==ทั้งหมด=="
        .AddItem "เยี่ยมชมศูนย์"
        .AddItem "ตรวจประเมินประจำปี"
        .AddItem "ลูกค้าติดตามดูทำลาย"
        .ListIndex = 0
    End With
End Sub
Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\PR01.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_CustVisitRpt as vw_CustVisit" & vbCrLf
        tmpSql = tmpSql & "WHERE visitdate>='" & dtFrom.ValueYMD & "' And visitdate<='" & dtTo.ValueYMD & "' and isCancel<>'Y' "
        
            If cboStatus.ListIndex > 0 Then
                    tmpSql = tmpSql & "And VisitType='" & cboStatus.Text & "' "
            End If
        
        If cboCall.ListIndex = 2 Then
            tmpSql = tmpSql & " AND Isinform='Y' "
        ElseIf cboCall.ListIndex = 1 Then
            tmpSql = tmpSql & " AND Isinform<>'Y' "
        End If
        
        If cboCustType.ListIndex = 1 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('C','S')"
        ElseIf cboCustType.ListIndex = 2 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('M','MS')"
        ElseIf cboCustType.ListIndex = 3 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('B','E')"
        End If
        
        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' " & vbCrLf
            End If
        End If

        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & " AND UnderSaleStaffID='" & curStaff_.ID & "'"
        End If
        tmpSql = tmpSql & vbCrLf & "ORDER BY visitdate,CompanyName,SaleStaffName"
'        .Formulas(0) = "PRStaff=" & Chr(34) & "ข้อมูลวันที่ " & " " & dtFrom.ValueDMY & " " & " ถึงวันที่ " & " " & dtTo.ValueDMY & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
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
