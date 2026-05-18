VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRptord06 
   Caption         =   "รายงาน Order ที่ยกเลิก"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   675
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2850
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2850
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   675
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2850
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2040
      Left            =   -15
      TabIndex        =   0
      Top             =   765
      Width           =   6045
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1125
         TabIndex        =   16
         Top             =   765
         Width           =   4335
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1215
         TabIndex        =   14
         Top             =   45
         Visible         =   0   'False
         Width           =   4365
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1110
         TabIndex        =   12
         Text            =   "cboCarType"
         Top             =   1170
         Width           =   4395
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "FrmRptord06.frx":0000
         Left            =   1125
         List            =   "FrmRptord06.frx":0002
         TabIndex        =   1
         Top             =   1530
         Width           =   1815
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
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
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3750
         TabIndex        =   3
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
         TabIndex        =   17
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทขนส่ง"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   15
         Top             =   90
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   13
         Top             =   1230
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   6
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   5
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   4
         Top             =   1545
         Width           =   930
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานการยกเลิก Order"
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
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2130
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "FrmRptord06.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานสรุปเบี้ยเลี้ยง (TRN16)"
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
      Index           =   6
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   2610
   End
End
Attribute VB_Name = "FrmRptord06"
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

Private Sub Form_Load()
     With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(M) ,(MS)"
        .AddItem "(B) (E)"
        .ListIndex = 0
    End With
    PopulateCarType cboCarType

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

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, dateEng As String  ', StaffID As String
Dim strCond As String
'    StaffID = ""
'    If Not curStaff_ Is Nothing Then
'        If Trim(curStaff_.ID) <> "" Then
'            StaffID = curStaff_.ID
'        End If
'    End If
    
    strCond = "WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'"
    If cboCustType.ListIndex = 1 Then
        strCond = strCond & " AND CustTypeFlag in ('C','S') "
    ElseIf cboCustType.ListIndex = 2 Then
        strCond = strCond & " AND CustTypeFlag in ('M','MS') "
   ElseIf cboCustType.ListIndex = 3 Then
        strCond = strCond & " AND CustTypeFlag in ('B','E') "
    End If
        If cboCarType.ListIndex > 0 Then
            strCond = strCond & " AND truckTypeDesc = '" & Trim(cboCarType.Text) & "' "
        End If

        If Trim(txtSale.Text) <> "" Then
            strCond = strCond & "AND UnderSaleStaffID='" & curStaff_.ID & "' " & vbCrLf
        End If

    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\rptord06.rpt"
        tmpSql = ""
        tmpSql = "select CustomerCode ,CompanyName ,WorkDate ,isCancel,cancelTime,JobNo ,SubTypeDesc ,CancelRemark ,cancelStaff,SaleName ,CustTypeFlag,TruckCount  from vw_printord06 " & vbCrLf
        tmpSql = tmpSql & "WHERE " & strCond
        tmpSql = tmpSql & vbCrLf & "  ORDER BY CancelRemark,workdate"
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
