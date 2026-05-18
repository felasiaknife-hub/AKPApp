VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4464.109
   ScaleMode       =   0  'User
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5985
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   900
         Width           =   4425
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1140
         TabIndex        =   11
         Text            =   "cboCarType"
         Top             =   1320
         Visible         =   0   'False
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3810
         TabIndex        =   2
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   1380
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ขนโดย"
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   4
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   3
         Top             =   420
         Width           =   795
      End
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "พิมพ์รายงานเมนิเฟสยังไม่วางบิล (ไม่ได้ของ)"
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
      Left            =   420
      TabIndex        =   8
      Top             =   0
      Width           =   3825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   119.706
      X2              =   5805.735
      Y1              =   779.757
      Y2              =   779.757
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   179.559
      X2              =   5805.735
      Y1              =   623.806
      Y2              =   623.806
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintTRN14.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim coltruckType_ As Collection
Dim curStaff_ As BWGCompanyStaff

Private Sub cmdPreview_Click()
PrintMain 0
End Sub

Private Sub cmdPrint_Click()
 PrintMain 0
End Sub

Private Sub Form_Load()
    PopulateCarType cboCarType
    PopulateCompany cboCompany, cboIndex
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = 1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\FIN10.rpt"
        tmpSql = "select WorkDate,JobNo,TimetableNo,DocNo,CustomerName,TruckTypeDesc," & vbCrLf
        tmpSql = tmpSql & " TripTransFee_Inc,VDTransFee,TransportName2,FailRemark " & vbCrLf
        tmpSql = tmpSql & " From vw_RptTrn14 " & vbCrLf
        tmpSql = tmpSql & " Where (WorkDate between '" & dtFrom.ValueYMD & "' and '" & dtTo.ValueYMD & "') " & vbCrLf
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        End If
        tmpSql = tmpSql & " Order by WorkDate,TimetableNo,CustomerName "
        .Formulas(0) = "Company=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(1) = "DateReport=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .WindowTitle = "รายงานการขนที่ไม่ได้ของ  "
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

