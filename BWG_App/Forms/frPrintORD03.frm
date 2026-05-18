VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintORD03 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F094"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rpt1 
      Left            =   180
      Top             =   2220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1950
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   90
      TabIndex        =   4
      Top             =   810
      Width           =   6045
      Begin VB.ComboBox cboCommissionType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   1905
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าคอมมิชชั่นของ"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3450
         TabIndex        =   5
         Top             =   660
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1950
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1950
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
      Caption         =   "พิมพ์รายงานค่าคอมมิชชั่น"
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
      Left            =   150
      TabIndex        =   3
      Top             =   60
      Width           =   2235
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintORD03.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintORD03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const rowPerBill As Integer = 37
Const itemPerBill As Integer = 25
Const reportRowPerBill As Integer = 9

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    'Me.CryReport.Connect = g_ConnectionStringdns
                     
    Me.rpt1.WindowLeft = 0: Me.rpt1.WindowWidth = 800
    Me.rpt1.WindowTop = 0: Me.rpt1.WindowHeight = 600
    Me.rpt1.Destination = crptToWindow
    
'    Me.rpt1.ParameterFields(0) = ""
'    Me.rpt1.ParameterFields(1) = ""
'    Me.rpt1.ParameterFields(2) = ""
'    Me.rpt1.ParameterFields(3) = ""
'    Me.rpt1.ParameterFields(4) = ""
'                'varPaidSeq = 1
'                 Me.CryReport.ParameterFields(1) = "PPaidq ;" & varPaidSeq & ";true"
'                ' Me.CryReport.ParameterFields(2) = "PReportName;" & strReportName & ";true"
'                ' Me.CryReport.ParameterFields(3) = "PReceiptNo; " & Trim(Me.txtCustomerNo.Text) & ";Ture"
'                ' Me.CryReport.ParameterFields(4) = "dtDateReturn; " & varDate & ";Ture"
                 
    Me.rpt1.ReportFileName = App.Path + "\Reports\ORD03.rpt"
    
    Me.rpt1.Action = 1
    Screen.MousePointer = 0
   
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

