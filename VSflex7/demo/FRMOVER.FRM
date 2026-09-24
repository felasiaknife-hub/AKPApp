VERSION 5.00
Begin VB.Form frmOverall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VideoSoft VSFlexGrid Pro 7.0"
   ClientHeight    =   7230
   ClientLeft      =   345
   ClientTop       =   870
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmOver.frx":0000
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   9180
      X2              =   6480
      Y1              =   3300
      Y2              =   3300
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6480
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   11280
      X2              =   11280
      Y1              =   3840
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   11280
      X2              =   6480
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6480
      X2              =   6480
      Y1              =   3840
      Y2              =   6960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   3180
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   5760
      X2              =   720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   6480
      Picture         =   "frmOver.frx":05FF
      Top             =   2160
      Width           =   2475
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   11280
      X2              =   6480
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   600
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   135
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   7
      Left            =   6840
      Picture         =   "frmOver.frx":25C8
      Top             =   4020
      Width           =   1875
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   6
      Left            =   6840
      Picture         =   "frmOver.frx":3D13
      Top             =   6540
      Width           =   1950
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   5
      Left            =   6840
      Picture         =   "frmOver.frx":548D
      Top             =   6120
      Width           =   3435
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   4
      Left            =   6840
      Picture         =   "frmOver.frx":7DAE
      Top             =   5700
      Width           =   3180
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   3
      Left            =   6840
      Picture         =   "frmOver.frx":A334
      Top             =   5280
      Width           =   2940
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   2
      Left            =   6840
      Picture         =   "frmOver.frx":C489
      Top             =   4860
      Width           =   4140
   End
   Begin VB.Image img 
      Height          =   300
      Index           =   1
      Left            =   6840
      Picture         =   "frmOver.frx":F2CE
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Image img 
      Height          =   1260
      Index           =   0
      Left            =   720
      Picture         =   "frmOver.frx":10F74
      Top             =   240
      Width           =   2235
   End
   Begin VB.Image imgLogo 
      Height          =   3780
      Left            =   300
      Top             =   3420
      Width           =   4380
   End
End
Attribute VB_Name = "frmOverall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    imgLogo = frmMenu!imgLogo
End Sub


