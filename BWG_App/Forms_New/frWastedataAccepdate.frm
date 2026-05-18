VERSION 5.00
Begin VB.Form frmWastedataAccepdate 
   Caption         =   "วันที่ตรวจรับ"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin WasteManagment.ctlDate ctlDate 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   300
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   435
      Left            =   2040
      TabIndex        =   2
      Top             =   780
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "วันที่ตรวจรับ :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   1095
   End
End
Attribute VB_Name = "frmWastedataAccepdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public wastedataID As String

Private Sub Command1_Click()
    frmWasteData.Accepdate = ctlDate.ValueYMD
    frmWasteData.setAccepDate ctlDate.ValueDMY
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ctlDate.ValueYMD = frmWasteData.Accepdate
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
