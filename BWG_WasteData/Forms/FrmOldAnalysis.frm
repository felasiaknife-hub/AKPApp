VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOldAnalysis 
   Caption         =   "ผลวิเคราะห์"
   ClientHeight    =   5355
   ClientLeft      =   4215
   ClientTop       =   5145
   ClientWidth     =   12570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   12570
   Begin VB.CommandButton CmdClose 
      Caption         =   "ปิด"
      Height          =   735
      Left            =   11250
      TabIndex        =   0
      Top             =   4590
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvAnalysis 
      Height          =   4170
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   7355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "วันที่วิเคราะห์"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่วิเคราะห์"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อ Waste"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ประเภท"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "รหัสลูกค้า"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "เซล"
         Object.Width           =   3528
      EndProperty
   End
   Begin Crystal.CrystalReport CRrpt 
      Left            =   5895
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "เลือกผลวิเคราะห์ (Double Click เพื่อเปิดรายงาน)"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   3360
   End
End
Attribute VB_Name = "FrmOldAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormType As Integer

Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub PrintRpt()
   Dim tSql As String
    tSql = "SELECT * FROM vw_wastedatalabrev as vw_WasteDataCRSearch where id =" & lvAnalysis.SelectedItem.Key & " "
    With CRrpt
        .Reset
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\WasteDataAnalysis.rpt"
        .WindowTitle = "ข้อมูลวิเคราะห์เวสต์เก่า"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
'        .SelectionFormula = "{vw_WasteDataCRSearch.WasteDataID} = '" & Replace(WasteDataID, "'", "") & "'"
        .SQLQuery = tSql
        .Destination = 0
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
Private Sub lvAnalysis_DblClick()
    PrintRpt
End Sub
