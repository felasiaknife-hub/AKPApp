VERSION 5.00
Begin VB.Form frmDateAnalysis 
   Caption         =   "วันที่วิเคราะห์"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4125
   LinkTopic       =   "From1"
   MaxButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   435
      Left            =   2040
      TabIndex        =   0
      Top             =   900
      Width           =   915
   End
   Begin WasteDataForm.ctlDate ctlDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "วันที่วิเคราะห์ :"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "frmDateAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmWasteData.AnalysisDate = ctlDate.ValueYMD
frmWasteData.SetAnalysisDate
Unload Me
End Sub
    
Private Sub Command2_Click()
    Unload Me
End Sub
