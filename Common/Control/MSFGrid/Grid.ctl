VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl Grid 
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   ScaleHeight     =   5250
   ScaleWidth      =   9255
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      _Version        =   393216
   End
End
Attribute VB_Name = "Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private clsFGEdit As clsFlexGridEdit

Private Sub UserControl_Initialize()
    Set clsFGEdit = New clsFlexGridEdit
    Set clsFGEdit.FlexGridControl = MSFlexGrid1
End Sub

Private Sub UserControl_Resize()
    MSFlexGrid1.Width = UserControl.Width
    MSFlexGrid1.Height = UserControl.Height
End Sub
