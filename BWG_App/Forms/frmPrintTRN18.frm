VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN18 
   Caption         =   "สรุปเบี้ยเลี้ยง TRN16"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6045
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   360
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
         TabIndex        =   7
         Top             =   420
         Width           =   495
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
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   675
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   675
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   915
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
      TabIndex        =   9
      Top             =   0
      Width           =   2610
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintTRN18.frx":0000
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
      TabIndex        =   8
      Top             =   0
      Width           =   2610
   End
End
Attribute VB_Name = "frmPrintTRN18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub PrintMain(printType As Integer)
    With rpt1
                .ReportFileName = App.Path & "\reports\TRN18.rpt"
             .SelectionFormula = "{vw_RptTRN17.WorkDate} in " & Chr(34) & dtFrom.ValueYMD & Chr(34) & " to " & Chr(34) & dtTo.ValueYMD & Chr(34)
        

        .Connect = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        '.Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(dtFrom.ValueDMY, "/", "/") & Chr(34)
         '.Formulas(2) = "ShowPrintDateTo=" & Chr(34) & FormatYMD_to_DMY(dtTo.ValueDMY, "/", "/") & Chr(34)
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub
