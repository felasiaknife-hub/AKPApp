VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   Caption         =   "VSFlexGrid Pro in HTML"
   ClientHeight    =   3615
   ClientLeft      =   2805
   ClientTop       =   2280
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4965
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2640
      Left            =   525
      TabIndex        =   0
      Top             =   390
      Width           =   4155
      ExtentX         =   7329
      ExtentY         =   4657
      ViewMode        =   1
      Offline         =   1
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    
    Me.WebBrowser1.Move 0, 0, ScaleWidth, ScaleHeight

End Sub


