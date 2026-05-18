VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN04 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Top             =   780
      Width           =   6045
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   4395
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   1140
         Width           =   4365
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
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
         TabIndex        =   1
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
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   810
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เจ้าหนี้"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   10
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   9
         Top             =   420
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   915
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
      Caption         =   "พิมพ์รายงานค่าขนส่ง,ค่าแรงคนงาน (TRN04)"
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
      Width           =   3900
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintTRN04.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN04"
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
    If Trim(VendorID) <> "" Then
        strCond = strCond & " And VendorID='" & VendorID & "' "
    End If
    
    With rpt1
        If Trim(txtVendor.Text) <> "" Then
                .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
                .ReportFileName = App.Path & "\reports\RPTTRN04.rpt"
        Else
                .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
                .ReportFileName = App.Path & "\reports\RPTTRN04-1.rpt"
        End If

        
         tmpSql = ""
        tmpSql = "select  *  from vw_RptTrn15_1 " & vbCrLf
        tmpSql = tmpSql & "WHERE " & strCond
        tmpSql = tmpSql & vbCrLf & "  ORDER BY  WorkDate,TimeTableNo,TrainRegisID ASC"
        .Formulas(0) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(dtFrom.ValueDMY, "/", "/") & Chr(34) & "ถึง" & Chr(34) & FormatYMD_to_DMY(dtTo.ValueDMY, "/", "/") & Chr(34)
        .Formulas(0) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)

        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBConnExc "Delete tbRptFIN08 WHERE SessionID='" & sessionID_ & "'"
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
    End With
End Sub





