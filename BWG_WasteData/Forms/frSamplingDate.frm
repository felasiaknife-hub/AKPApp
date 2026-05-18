VERSION 5.00
Begin VB.Form frmSamplingDate 
   Caption         =   "วันที่เก็บตัวอย่าง :"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Height          =   435
      Left            =   1200
      TabIndex        =   1
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   435
      Left            =   2400
      TabIndex        =   0
      Top             =   900
      Width           =   915
   End
   Begin WasteDataForm.ctlDate ctlDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "วันที่ตรวจรับ :"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "frmSamplingDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    frmWasteData.SamplingDate = ctlDate.ValueYMD
    frmWasteData.SetSamplingDate
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
