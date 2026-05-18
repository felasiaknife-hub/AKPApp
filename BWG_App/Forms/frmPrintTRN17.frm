VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN17 
   Caption         =   "รายงานการวาง Box"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   675
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   675
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   6045
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   4395
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
         Caption         =   "ในนามบริษัท"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   810
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   4
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานการวาง Box (TRN17)"
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
      TabIndex        =   7
      Top             =   0
      Width           =   2520
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintTRN17.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim curVendor_ As BWGTransportVendor
'Dim sessionID_ As String
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
'     With cboCustType
'        .AddItem "== ประเภทลูกค้า =="
'        .AddItem "(C) ,(s)"
'        .AddItem "(F) RDF Waste"
'        .AddItem "(L) (p)"
'        .AddItem "(T) (k)"
'        .AddItem "(G) (e)"
'        .ListIndex = 0
'    End With
'    sessionID_ = GetGUID
    PopulateCompany cboCompany, cboIndex
'    PopulateVenderContact cboDrivername
End Sub
Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, dateEng As String, VendorID As String
Dim strCond As String
'    strCond = "WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'"
'    If cboCompany.ListCount > 1 Then
'        If cboCompany.ListIndex > 0 Then
'
'            strCond = strCond & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
'        Else
'
'        End If
'
'    Else
'        If cboCompany.ListIndex >= 0 Then
'            strCond = strCond & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
'        End If
'    End If
     With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\TRN17.rpt"
        tmpSql = ""
        If cboCompany.ListIndex >= 0 Then
            tmpSql = tmpSql & " EXEC sp_TrnBoxWork " & "'" & dtFrom.ValueYMD & "'," & "'" & dtTo.ValueYMD & "'," & "'" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
        Else
            tmpSql = tmpSql & " EXEC sp_TrnBoxWork " & "'" & dtFrom.ValueYMD & "'," & "'" & dtTo.ValueYMD & "'," & "''" & vbCrLf
        End If
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(dtFrom.ValueDMY, "/", "/") & Chr(34)
         .Formulas(2) = "ShowPrintDateTo=" & Chr(34) & FormatYMD_to_DMY(dtTo.ValueDMY, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub




