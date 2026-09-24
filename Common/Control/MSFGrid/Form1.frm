VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   4320
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2990
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Grid1
    MSFlexGrid1.TextMatrix(0, 0) = "TEST"
    MSFlexGrid1
End Sub

