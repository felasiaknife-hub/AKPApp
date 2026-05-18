VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComProcess 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ประมวลผลค่าคอมมิชชั่น - F015"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "-"
      Height          =   195
      Left            =   450
      TabIndex        =   9
      Top             =   480
      Width           =   405
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   195
      Left            =   30
      TabIndex        =   8
      Top             =   480
      Width           =   405
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ปิด"
      Height          =   435
      Left            =   10290
      TabIndex        =   7
      Top             =   90
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "คำนวณค่าคอม"
      Height          =   435
      Left            =   8910
      TabIndex        =   6
      Top             =   90
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "แสดงใบ Menifest"
      Height          =   435
      Left            =   7320
      TabIndex        =   5
      Top             =   90
      Width           =   1545
   End
   Begin WasteManagment.ctlDate ctlDate1 
      Height          =   345
      Left            =   2820
      TabIndex        =   1
      Top             =   150
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate ctlDate2 
      Height          =   345
      Left            =   5400
      TabIndex        =   3
      Top             =   150
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   8145
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   14367
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Menifest No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ลูกค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อกากของเสีย"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันทีขน"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "วันที่ชำระเงิน"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ยอดที่ชำระ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ค่าใช้จ่าย"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "% หักยอดขาย"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ยอดที่นำมาคิดค่าคอม"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   0
      Left            =   4740
      TabIndex        =   2
      Top             =   210
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menifest ที่มีการชำระเงินในช่วงวันที่ :"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   2625
   End
End
Attribute VB_Name = "frmComProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    lvResult.Width = Me.ScaleWidth - 60
    lvResult.Height = Me.ScaleHeight - lvResult.Top
End Sub
