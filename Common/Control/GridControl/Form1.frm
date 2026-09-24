VERSION 5.00
Object = "{6AB4F7FF-6DFD-4F67-A05A-EF7C0949F8C6}#3.0#0"; "GridCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin GridCtl.Grid Grid1 
      Height          =   5655
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Grid1.get
End Sub
