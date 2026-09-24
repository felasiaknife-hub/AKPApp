VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmfreez 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frozen Rows and Columns"
   ClientHeight    =   6045
   ClientLeft      =   705
   ClientTop       =   1545
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFrozenRows 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frozen Rows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox chkFrozenCols 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frozen Columns"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4155
      Left            =   240
      TabIndex        =   0
      Top             =   1740
      Width           =   9975
      _cx             =   17595
      _cy             =   7329
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   11
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmfreez.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   12640511
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Frozen and then scroll the grid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   420
      X2              =   6720
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   300
      Shape           =   3  'Circle
      Top             =   900
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   420
      Picture         =   "frmfreez.frx":04AF
      Top             =   180
      Width           =   5640
   End
End
Attribute VB_Name = "frmfreez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()

End Sub

Private Sub chkFrozenCols_Click()
    
    If chkFrozenCols Then
        fg.FrozenCols = 2
    Else
        fg.FrozenCols = 0
    End If
    
End Sub


Private Sub chkFrozenRows_Click()
    
    ' Set Frozen Rows
    If chkFrozenRows Then
        fg.FrozenRows = 3
    Else
        fg.FrozenRows = 0
    End If
    
End Sub

Private Sub Form_Load()
    Dim r&, c&
    With fg
        .MergeCells = flexMergeFree
        .MergeRow(-1) = True
        .MergeCol(-1) = True
        .Select 0, 3, 0, 10
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3, 0, 6) = "1998"
        .Cell(flexcpText, 0, 7, 0, 10) = "1999"
        
    
    End With
End Sub

