VERSION 5.00
Begin VB.Form frmGroup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VSFlexGrid Pro 6.0 - Grouping Demo"
   ClientHeight    =   6255
   ClientLeft      =   555
   ClientTop       =   1470
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin FlexDemo.VSFlexGroup VSFlexGroup1 
      Height          =   4635
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8176
   End
   Begin VB.Image imgItem 
      Height          =   225
      Left            =   5280
      Picture         =   "frmGroup.frx":0000
      Top             =   300
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgFolder 
      Height          =   225
      Left            =   5280
      Picture         =   "frmGroup.frx":00FA
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   420
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   540
      X2              =   3300
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   660
      Picture         =   "frmGroup.frx":01F4
      Top             =   240
      Width           =   1845
   End
End
Attribute VB_Name = "frmGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    ' fill control
    Dim fg As VSFlexGrid
    Set fg = VSFlexGroup1.FlexGrid
    
    ' set dimensions
    fg.Cols = 4
    fg.Rows = 100
    
    ' set data
    FillColumn fg, 0, "Categories", "Business|Personal|Engineering|Computing|Conferences|University"
    FillColumn fg, 1, "Subject", "Proposals|Demos|Social|Trips|Projects"
    FillColumn fg, 2, "Received", "1/1|2/2|3/3|4/4|5/5|6/6|7/7|8/8|9/9|10/10|11/11|December 12,2007"
    FillColumn fg, 3, "Flag Status", "Important|Junk|Handled|Backburner"
    
    ' some formatting
    fg.ColDataType(2) = flexDTDate
    fg.ColFormat(2) = "Medium Date"
    
    VSFlexGroup1.Update
    
End Sub


Sub FillColumn(fg As VSFlexGrid, ByVal c As Long, ByVal t As String, ByVal s As String)
    
    Dim r&, i&, cnt&
    ReDim lst(0) As String
    
    ' build list
    cnt = 0
    i = InStr(s, "|")
    While i > 0
        lst(cnt) = left(s, i - 1)
        s = Mid(s, i + 1)
        cnt = cnt + 1
        ReDim Preserve lst(cnt) As String
        i = InStr(s, "|")
    Wend
    lst(cnt) = s
    
    ' set values
    fg.Cell(flexcpText, 0, c) = t
    For r = fg.FixedRows To fg.Rows - 1
        i = (Rnd() * 1000) Mod (cnt + 1)
        fg.Cell(flexcpText, r, c) = lst(i)
    Next
    
    ' do an autosize
    fg.AutoSize c, , , 300
End Sub






