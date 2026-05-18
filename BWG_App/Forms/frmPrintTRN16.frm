VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintTRN16 
   Caption         =   "สรุปเบี้ยเลี้ยง TRN16"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6045
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   1140
         TabIndex        =   22
         Top             =   2040
         Width           =   1815
      End
      Begin VB.OptionButton Opt2 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   2760
         Width           =   255
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H80000005&
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   2760
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   750
         Width           =   4395
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1140
         Width           =   4365
      End
      Begin VB.ComboBox cboDrivername 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   1560
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รายงานแบบแสดงรายละเอียด"
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   20
         Top             =   2820
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รายงานแบบสรุป"
         Height          =   195
         Index           =   8
         Left            =   1440
         TabIndex        =   19
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   13
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ในนามบริษัท"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   810
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทขนส่ง"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   10
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขับ"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   1620
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   675
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   675
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บริษัทขนส่ง"
      Height          =   195
      Index           =   7
      Left            =   1560
      TabIndex        =   18
      Top             =   3000
      Width           =   1275
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
      TabIndex        =   15
      Top             =   0
      Width           =   2610
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintTRN16.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานสรุปเบี้ยเลี้ยง (TRN15)"
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
      TabIndex        =   14
      Top             =   0
      Width           =   2610
   End
End
Attribute VB_Name = "frmPrintTRN16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curVendor_ As BWGTransportVendor
Dim sessionID_ As String


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
 With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
     With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(F) RDF Waste"
        .AddItem "(L) (p)"
        .AddItem "(T) (k)"
        .AddItem "(G) (e)"
        .AddItem "(R)"
        .ListIndex = 0
    End With
    sessionID_ = GetGUID
    PopulateCompany cboCompany, cboIndex
    PopulateVenderContact cboDrivername
End Sub
Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, dateEng As String, VendorID As String
Dim strCond As String
    VendorID = ""
    If Not curVendor_ Is Nothing Then
        If Trim(curVendor_.ID) <> "" Then
            VendorID = curVendor_.ID
        End If
    End If
    
    strCond = "WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'"
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            strCond = strCond & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
        End If
    Else
        If cboCompany.ListIndex >= 0 Then
            strCond = strCond & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
        End If
    End If
    
    If Trim(VendorID) <> "" Then
        strCond = strCond & " And VendorID='" & VendorID & "' "
    End If
    If cboDrivername.ListIndex <> 0 Then
        strCond = strCond & " And DriverName = '" & cboDrivername.Text & "' "
    End If
    If Opt1.Value = True Then
    If cboCustType.ListIndex = 1 Then
        strCond = strCond & " AND CustTypeFlag in ('C','S') "
    ElseIf cboCustType.ListIndex = 2 Then
        strCond = strCond & " AND CustTypeFlag = 'F' "
   ElseIf cboCustType.ListIndex = 3 Then
        strCond = strCond & " AND CustTypeFlag in ('L','P') "
    ElseIf cboCustType.ListIndex = 4 Then
        strCond = strCond & " AND CustTypeFlag in ('T','K') "
    ElseIf cboCustType.ListIndex = 5 Then
        strCond = strCond & " AND CustTypeFlag in ('G','E') "
    ElseIf cboCustType.ListIndex = 6 Then
        strCond = strCond & " AND CustTypeFlag in ('R') "
    End If
    With Rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\RPTTRN16.rpt"
        tmpSql = ""
        tmpSql = "select *  from vw_RptTrn15_2 " & vbCrLf
        tmpSql = tmpSql & "WHERE " & strCond
        tmpSql = tmpSql & vbCrLf & "  ORDER BY  DriverName"
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(dtFrom.ValueDMY, "/", "/") & Chr(34)
         .Formulas(2) = "ShowPrintDateTo=" & Chr(34) & FormatYMD_to_DMY(dtTo.ValueDMY, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
Else
    
    With Rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\RPTTRN16_1.rpt"
        tmpSql = ""
        tmpSql = "select  *  from vw_RptTrn15_1 " & vbCrLf
        tmpSql = tmpSql & "WHERE " & strCond
        tmpSql = tmpSql & vbCrLf & "  ORDER BY  WorkDate,TimeTableNo"
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(dtFrom.ValueDMY, "/", "/") & Chr(34)
         .Formulas(2) = "ShowPrintDateTo=" & Chr(34) & FormatYMD_to_DMY(dtTo.ValueDMY, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End If
End Sub





Private Sub txtVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtVendor_Validate False
End Sub
Private Sub txtVendor_Validate(Cancel As Boolean)
Dim selStr As String
    If curVendor_ Is Nothing Then Set curVendor_ = New BWGTransportVendor
    selStr = Trim(curVendor_.NameTH)
    If Trim(txtVendor.Text) = "" And Trim(selStr) <> Trim(txtVendor.Text) Then
        Set curVendor_ = Nothing
        Set curVendor_ = New BWGTransportVendor
    Else
        If Trim(selStr) <> Trim(txtVendor.Text) Then
            Set curVendor_ = Nothing
            Set curVendor_ = New BWGTransportVendor
            Load frmTransportVenderSearch
            With frmTransportVenderSearch
                .getCriteria (txtVendor.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curVendor_ Is Nothing Then
                txtVendor.Text = ""
            Else
                If curVendor_.ID = "" Then
                    txtVendor.Text = ""
                End If
            End If
        End If
    End If
End Sub
Public Sub SetDataTransportVender(selVendor As BWGTransportVendor, Optional isAssigned As Boolean = True)
    Set curVendor_ = Nothing
    Set curVendor_ = selVendor
    If curVendor_ Is Nothing Then Exit Sub
    With curVendor_
        txtVendor.Text = .NameTH
        PopulateVenderContact cboDrivername, " VendorID = '" & curVendor_.ID & "' "
    End With
End Sub

