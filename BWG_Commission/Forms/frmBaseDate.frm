VERSION 5.00
Begin VB.Form frmBaseDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "กำหนดวันที่ตั้งยอด"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin Commission.ctlDate ctlDate1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "กำหนดเดือนที่ตั้งยอด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   240
      Width           =   1755
   End
End
Attribute VB_Name = "frmBaseDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmExtCom2.SetBaseDate ctlDate1.ValueYMD, False
    Unload Me
End Sub


Private Sub cmdOK_Click()
    frmExtCom2.SetBaseDate ctlDate1.ValueYMD, True
    Unload Me
End Sub

Private Sub Form_Load()
    ctlDate1.DayValue = "1"
End Sub


