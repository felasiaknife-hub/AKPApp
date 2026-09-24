VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsFlex7l.ocx"
Begin VB.Form frmPalette 
   BorderStyle     =   0  'None
   Caption         =   "Color Picker"
   ClientHeight    =   1965
   ClientLeft      =   4530
   ClientTop       =   5190
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VSFlex7LCtl.VSFlexGrid fgSystem 
      Height          =   1605
      Left            =   2250
      TabIndex        =   1
      Top             =   300
      Width           =   3090
      _cx             =   5450
      _cy             =   2822
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Editable        =   0
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
   End
   Begin VSFlex7LCtl.VSFlexGrid fgPalette 
      Height          =   1605
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   2130
      _cx             =   3757
      _cy             =   2822
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Rows            =   6
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
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
      Editable        =   0
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
   End
   Begin VB.Shape Shape1 
      Height          =   1965
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Colors"
      Height          =   195
      Left            =   2250
      TabIndex        =   3
      Top             =   75
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palette"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   510
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fgPalette_Click()

    ' return color selected
    With fgPalette
        Tag = .Cell(flexcpBackColor, .Row, .Col)
    End With
    Me.Hide
    
End Sub

Private Sub fgSystem_Click()
    
    ' return color selected
    With fgSystem
        Tag = .Cell(flexcpBackColor, .Row, 0)
    End With
    Me.Hide

End Sub

Private Sub Form_Deactivate()

    ' go away when deactivated
    Me.Hide
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Me.Hide
End Sub

Private Sub Form_Load()
    With fgSystem
    
        ' form wants keys
        KeyPreview = True
        
        ' initialize grid
        .Rows = 0
        .Cols = 2
        .FixedRows = 0
        .FixedCols = 0
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarVertical
        .AllowSelection = False
        .HighLight = flexHighlightNever
        .ColWidth(0) = 300
        .GridColor = .BackColor
        .GridLineWidth = 3
        
        ' populate grid
        .AddItem vbTab & "Scroll Bars"
        .AddItem vbTab & "Desktop"
        .AddItem vbTab & "Active Title Bar"
        .AddItem vbTab & "Inactive Title Bar"
        .AddItem vbTab & "Menu Bar"
        .AddItem vbTab & "Window Background"
        .AddItem vbTab & "Window Frame"
        .AddItem vbTab & "Menu Text"
        .AddItem vbTab & "Window Text"
        .AddItem vbTab & "Active Title Bar Text"
        .AddItem vbTab & "Active Border"
        .AddItem vbTab & "Inactive Border"
        .AddItem vbTab & "Application Workspace"
        .AddItem vbTab & "HighLight"
        .AddItem vbTab & "HighLight Text"
        .AddItem vbTab & "Button Face"
        .AddItem vbTab & "Button Shadow"
        .AddItem vbTab & "Disabled Text"
        .AddItem vbTab & "Button Text"
        .AddItem vbTab & "Inactive Title Bar Text"
        .AddItem vbTab & "Button HighLight"
        .AddItem vbTab & "Button Dark Shadow"
        .AddItem vbTab & "Button Light Shadow"
        .AddItem vbTab & "ToolTip Text"
        .AddItem vbTab & "ToolTip"
        
        ' show color preview
        Dim r%
        For r = 0 To .Rows - 1
            .Cell(flexcpBackColor, r, 0) = &H80000000 + r
        Next
    
    End With
    
    With fgPalette
    
        ' initialize grid
        .Rows = 6
        .Cols = 8
        .FixedRows = 0
        .FixedCols = 0
        .ExtendLastCol = True
        .ScrollBars = flexScrollBarNone
        .AllowSelection = False
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .RowHeight(-1) = .ClientHeight / .Rows
        .ColWidth(-1) = .ClientWidth / .Cols
        .RowHeight(0) = .RowHeight(.Rows - 1) + Screen.TwipsPerPixelY
        .ColWidth(0) = .ColWidth(.Cols - 1) + Screen.TwipsPerPixelX
        
        ' set back colors
        ' gray ----------------------------
        .Cell(flexcpBackColor, 0, 0) = 1
        .Cell(flexcpBackColor, 1, 0) = 15198183
        .Cell(flexcpBackColor, 2, 0) = 13027014
        .Cell(flexcpBackColor, 3, 0) = 8684676
        .Cell(flexcpBackColor, 4, 0) = 4342338
        .Cell(flexcpBackColor, 5, 0) = 1
        ' red ----------------------------
        .Cell(flexcpBackColor, 0, 1) = 13027071
        .Cell(flexcpBackColor, 1, 1) = 8684799
        .Cell(flexcpBackColor, 2, 1) = 255
        .Cell(flexcpBackColor, 3, 1) = 198
        .Cell(flexcpBackColor, 4, 1) = 132
        .Cell(flexcpBackColor, 5, 1) = 66
        ' orange ----------------------------
        .Cell(flexcpBackColor, 0, 2) = 13035519
        .Cell(flexcpBackColor, 1, 2) = 8701695
        .Cell(flexcpBackColor, 2, 2) = 34047
        .Cell(flexcpBackColor, 3, 2) = 17094
        .Cell(flexcpBackColor, 4, 2) = 17028
        .Cell(flexcpBackColor, 5, 2) = 4342404
        ' yellow ----------------------------
        .Cell(flexcpBackColor, 0, 3) = 13041663
        .Cell(flexcpBackColor, 1, 3) = 8716287
        .Cell(flexcpBackColor, 2, 3) = 65535
        .Cell(flexcpBackColor, 3, 3) = 50886
        .Cell(flexcpBackColor, 4, 3) = 33924
        .Cell(flexcpBackColor, 5, 3) = 16962
        ' green ----------------------------
        .Cell(flexcpBackColor, 0, 4) = 13041606
        .Cell(flexcpBackColor, 1, 4) = 8716164
        .Cell(flexcpBackColor, 2, 4) = 65280
        .Cell(flexcpBackColor, 3, 4) = 50688
        .Cell(flexcpBackColor, 4, 4) = 33792
        .Cell(flexcpBackColor, 5, 4) = 16896
        ' cyan ----------------------------
        .Cell(flexcpBackColor, 0, 5) = 16777158
        .Cell(flexcpBackColor, 1, 5) = 16777092
        .Cell(flexcpBackColor, 2, 5) = 16776960
        .Cell(flexcpBackColor, 3, 5) = 13026816
        .Cell(flexcpBackColor, 4, 5) = 8684544
        .Cell(flexcpBackColor, 5, 5) = 4342272
        ' blue ----------------------------
        .Cell(flexcpBackColor, 0, 6) = 16762566
        .Cell(flexcpBackColor, 1, 6) = 16745604
        .Cell(flexcpBackColor, 2, 6) = 16711680
        .Cell(flexcpBackColor, 3, 6) = 12976128
        .Cell(flexcpBackColor, 4, 6) = 8650752
        .Cell(flexcpBackColor, 5, 6) = 4325376
        ' magenta ----------------------------
        .Cell(flexcpBackColor, 0, 7) = 16762623
        .Cell(flexcpBackColor, 1, 7) = 16745727
        .Cell(flexcpBackColor, 2, 7) = 16711935
        .Cell(flexcpBackColor, 3, 7) = 12976326
        .Cell(flexcpBackColor, 4, 7) = 8650884
        .Cell(flexcpBackColor, 5, 7) = 4325442
        
    End With
End Sub
