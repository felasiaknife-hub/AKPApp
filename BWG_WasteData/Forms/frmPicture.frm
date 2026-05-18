VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDel 
      Height          =   1155
      Left            =   2460
      Picture         =   "frmPicture.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   1035
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox picEdit 
      Height          =   1155
      Left            =   1320
      Picture         =   "frmPicture.frx":04DA
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   9
      Top             =   2520
      Width           =   1035
   End
   Begin VB.PictureBox picAdd 
      Height          =   1155
      Left            =   60
      Picture         =   "frmPicture.frx":066F
      ScaleHeight     =   1095
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.PictureBox picApprove 
      Height          =   1155
      Left            =   3600
      Picture         =   "frmPicture.frx":0B7C
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   7
      Top             =   1170
      Width           =   1035
   End
   Begin VB.PictureBox picExit 
      Height          =   1065
      Left            =   60
      Picture         =   "frmPicture.frx":10CB
      ScaleHeight     =   1005
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   0
      Width           =   1035
   End
   Begin VB.PictureBox picSave 
      Height          =   1065
      Left            =   1200
      Picture         =   "frmPicture.frx":1647
      ScaleHeight     =   1005
      ScaleWidth      =   975
      TabIndex        =   5
      Top             =   0
      Width           =   1035
   End
   Begin VB.PictureBox picSearch 
      Height          =   1065
      Left            =   2370
      Picture         =   "frmPicture.frx":180A
      ScaleHeight     =   1005
      ScaleWidth      =   975
      TabIndex        =   4
      Top             =   0
      Width           =   1035
   End
   Begin VB.PictureBox picNew 
      Height          =   1065
      Left            =   3540
      Picture         =   "frmPicture.frx":1B48
      ScaleHeight     =   1005
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   0
      Width           =   1035
   End
   Begin VB.PictureBox picPrint 
      Height          =   1065
      Left            =   0
      Picture         =   "frmPicture.frx":1F05
      ScaleHeight     =   1005
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   1230
      Width           =   1035
   End
   Begin VB.PictureBox picDelete 
      Height          =   1155
      Left            =   2460
      Picture         =   "frmPicture.frx":24C2
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   1170
      Width           =   1035
   End
   Begin VB.PictureBox picPreview 
      Height          =   1155
      Left            =   1320
      Picture         =   "frmPicture.frx":2D8C
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   1170
      Width           =   1035
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

