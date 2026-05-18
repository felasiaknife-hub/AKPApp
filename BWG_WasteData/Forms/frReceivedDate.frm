VERSION 5.00
Begin VB.Form frmReceivedDate 
   Caption         =   "วันที่รับตัวอย่างกากฯ"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Height          =   435
      Left            =   1080
      TabIndex        =   1
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   435
      Left            =   2280
      TabIndex        =   0
      Top             =   780
      Width           =   915
   End
   Begin WasteDataForm.ctlDate ctlDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "วันที่รับตัวอย่าง :"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   300
      Width           =   1095
   End
End
Attribute VB_Name = "frmReceivedDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    frmWasteData.ReceivedDate = ctlDate.ValueYMD
    frmWasteData.SetReceivedDate
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
