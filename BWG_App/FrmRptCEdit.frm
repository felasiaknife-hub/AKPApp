VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRptCEdit 
   Caption         =   "เลือกข้อมูลที่ต้องการ"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6045
      Begin VB.ComboBox cboCon 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   255
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   855
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
         TabIndex        =   5
         Top             =   855
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   915
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ในนามบริษัท"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่"
         Height          =   195
         Index           =   3
         Left            =   675
         TabIndex        =   6
         Top             =   915
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   915
   End
End
Attribute VB_Name = "FrmRptCEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub Form_Load()
         With cboCon
        .AddItem "== เลือกภาชนะ =="
        .AddItem "ถังปากกว้างเหล็ก 200 ลิตร"
        .AddItem "ถังปากจู๋เหล็ก 200 ลิตร"
        .AddItem "แล็ค"
        .AddItem "Bigbag"
        .AddItem "เบ้า"
        .AddItem "ถุงดำ"
        .AddItem "กล่องกระดาษ"
        .AddItem "เบ้าตัดปาก"
        .ListIndex = 0
    End With

End Sub
Private Sub PrintMain(printType As Integer)
Dim tmpSql

'    strCond = "WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'"
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\RptContaineredit.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM Vw_Containeredit  " & vbCrLf
        tmpSql = tmpSql & " WHERE  (edate >= '" & dtFrom.ValueYMD & "'  and edate <= '" & dtTo.ValueYMD & "')"
        If cboCon.ListIndex <> 0 Then
        tmpSql = tmpSql & " and  conid='" & cboCon.ListIndex & "' "
        End If
        .Formulas(1) = "daterange=" & Chr(34) & "จากวันที่ " & (dtFrom.ValueDMY) & " ถึงวันที่ " & (dtTo.ValueDMY) & Chr(34)
        .Formulas(2) = "ShowPrintDate=" & Chr(34) & "วันที่พิมพ์ " & Now & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
