VERSION 5.00
Begin VB.Form frmPicture 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDone 
      Height          =   615
      Left            =   840
      Picture         =   "frmPicture.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   2340
      Width           =   615
   End
   Begin VB.PictureBox picOpen 
      Height          =   615
      Left            =   60
      Picture         =   "frmPicture.frx":0E42
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   2340
      Width           =   615
   End
   Begin VB.PictureBox picBox 
      Height          =   615
      Left            =   2400
      Picture         =   "frmPicture.frx":5B9C
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox picDel 
      Height          =   615
      Left            =   1620
      Picture         =   "frmPicture.frx":6466
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox picEdit 
      Height          =   615
      Left            =   840
      Picture         =   "frmPicture.frx":6940
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox picAdd 
      Height          =   615
      Left            =   60
      Picture         =   "frmPicture.frx":6AD5
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox picApprove 
      Height          =   615
      Left            =   2400
      Picture         =   "frmPicture.frx":6FE2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   780
      Width           =   615
   End
   Begin VB.PictureBox picExit 
      Height          =   615
      Left            =   60
      Picture         =   "frmPicture.frx":7531
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   840
      Picture         =   "frmPicture.frx":7AAD
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picSearch 
      Height          =   615
      Left            =   1620
      Picture         =   "frmPicture.frx":7C70
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picNew 
      Height          =   615
      Left            =   2400
      Picture         =   "frmPicture.frx":7FAE
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picPrint 
      Height          =   615
      Left            =   60
      Picture         =   "frmPicture.frx":836B
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   780
      Width           =   615
   End
   Begin VB.PictureBox picDelete 
      Height          =   615
      Left            =   1620
      Picture         =   "frmPicture.frx":8928
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   780
      Width           =   615
   End
   Begin VB.PictureBox picPreview 
      Height          =   615
      Left            =   840
      Picture         =   "frmPicture.frx":91F2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   780
      Width           =   615
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

