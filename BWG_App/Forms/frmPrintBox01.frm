VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintBox01 
   Caption         =   "รายงานการวาง Box ลูกค้า"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   675
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   675
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   6045
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
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
         Caption         =   "วันที่ว่าง"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   4
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   3
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
      TabIndex        =   5
      Top             =   0
      Width           =   2520
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintBox01.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintBox01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPreview_Click()
 PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub
Private Sub PrintMain(printType As Integer)
Dim tmpSql

'    strCond = "WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'"
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\RPTBox.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_BoxCustomer  " & vbCrLf
        tmpSql = tmpSql & " WHERE  (LUGBOXFrom >= '" & dtFrom.ValueYMD & "'  and LUGBOXFrom <= '" & dtTo.ValueYMD & "')" & vbCrLf
        tmpSql = tmpSql & " or (ROLBOXFrom >= '" & dtFrom.ValueYMD & "' and ROLBOXFrom <= '" & dtTo.ValueYMD & "')  " & vbCrLf
        tmpSql = tmpSql & " or (OneTFrom >= '" & dtFrom.ValueYMD & "'  and OneTFrom <= '" & dtTo.ValueYMD & "')  " & vbCrLf
        tmpSql = tmpSql & " or (TwoTFrom >= '" & dtFrom.ValueYMD & "'  and TwoTFrom <= '" & dtTo.ValueYMD & "') " & vbCrLf
        tmpSql = tmpSql & " or (TruckBOXFrom >= '" & dtFrom.ValueYMD & "'  and TruckBOXFrom <= '" & dtTo.ValueYMD & "') " & vbCrLf
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(dtFrom.ValueDMY, "/", "/") & Chr(34)
        .Formulas(2) = "ShowPrintDateTo=" & Chr(34) & FormatYMD_to_DMY(dtTo.ValueDMY, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

