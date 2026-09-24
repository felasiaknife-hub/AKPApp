VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid Binding"
   ClientHeight    =   6240
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7560
   Begin VB.CommandButton cmdLoadGrid 
      Caption         =   "Load from Flex"
      Height          =   300
      Left            =   5460
      TabIndex        =   5
      Top             =   60
      Width           =   1600
   End
   Begin VB.CommandButton cmdLoadArr 
      Caption         =   "Load from Array"
      Height          =   300
      Left            =   1740
      TabIndex        =   4
      Top             =   60
      Width           =   1600
   End
   Begin VB.CommandButton cmdBindGrid 
      Caption         =   "Bind to Flex"
      Height          =   300
      Left            =   3780
      TabIndex        =   3
      Top             =   60
      Width           =   1600
   End
   Begin VB.CommandButton cmdBindArr 
      Caption         =   "Bind to Array"
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1600
   End
   Begin VSFlex7LCtl.VSFlexGrid fg1 
      Height          =   2340
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   3765
      _cx             =   6641
      _cy             =   4128
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorAlternate=   -2147483643
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FormatCells     =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid fg2 
      Height          =   2580
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   3765
      _cx             =   6641
      _cy             =   4551
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColorAlternate=   -2147483643
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FormatCells     =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lbl 
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VarArr(4, 20)

Private Sub cmdBindArr_Click()
    
    ' bind to array
    fg1.BindToArray VarArr
    
    ' hide bottom grid
    fg2.Visible = False
    Form_Resize
    
    ' explain
    lbl = "The grid is bound to a global Variant Array. Changes made to the grid are reflected in the array, and changes made to the array are visible on the grid."

End Sub

Private Sub cmdBindGrid_Click()

    ' bind to bottom grid
    fg1.BindToArray fg2
    
    ' show bottom grid
    fg2.Visible = True
    Form_Resize
    
    ' explain
    lbl = "The top grid is bound to the bottom grid. Changes made to the either grid are reflected on the other grid."

End Sub

Private Sub cmdLoadArr_Click()

    ' load data from array
    fg1.LoadArray VarArr
    
    ' hide bottom grid
    fg2.Visible = False
    Form_Resize

    ' explain
    lbl = "The grid data was loaded from a global Variant Array. The grid and the array are now independent."

End Sub

Private Sub cmdLoadGrid_Click()
    
    ' load data from bottom grid
    fg1.LoadArray fg2
    
    ' show bottom grid
    fg2.Visible = True
    Form_Resize

    ' explain
    lbl = "The top grid data was loaded from the bottom grid. The grids are now independent."

End Sub

Private Sub Form_Load()

    ' initialize grids
    fg1.Editable = flexEDKbdMouse
    fg1.AllowUserResizing = flexResizeBoth
    fg2.Editable = flexEDNone
    fg2.AllowUserResizing = flexResizeBoth
    
    ' put some dummy data into second flex
    Dim r&, c&
    With fg2
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .TextMatrix(r, c) = "r" & r & "c" & c
            Next
        Next
    End With
    
    ' put some dummy data into Variant array
    For r = 0 To UBound(VarArr, 2)
        For c = 0 To UBound(VarArr, 1)
            VarArr(c, r) = "Arr " & r & ", " & c
        Next
    Next
    
    ' start bound to array
    cmdBindArr_Click
    
End Sub

Private Sub Form_Resize()

    ' just in case the form gets really small
    On Error Resume Next
    
    ' resize top grid
    On Error Resume Next
    Dim h!
    h = ScaleHeight - fg1.Top - fg1.Left
    If fg2.Visible Then h = h / 2
    fg1.Move fg1.Left, fg1.Top, ScaleWidth - 2 * fg1.Left, h
    
    ' resize bottom grid
    If fg2.Visible Then
        fg2.Move fg1.Left, fg1.Top + fg1.Height + fg1.Left, fg1.Width, fg1.Height - fg1.Left
    End If
    
    ' resize label
    lbl.Width = fg1.Width
End Sub
