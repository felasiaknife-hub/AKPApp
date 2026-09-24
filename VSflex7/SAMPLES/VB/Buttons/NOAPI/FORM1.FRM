VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlex7: Custom buttons (always visible)"
   ClientHeight    =   3645
   ClientLeft      =   2775
   ClientTop       =   1950
   ClientWidth     =   6585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   6585
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5655
      _cx             =   9975
      _cy             =   4577
      _ConvInfo       =   -1
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Image imgBtnDn 
      Height          =   210
      Left            =   6060
      Picture         =   "Form1.frx":0442
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBtnUp 
      Height          =   210
      Left            =   5700
      Picture         =   "Form1.frx":056A
      Top             =   180
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' only interesetd in left button
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = fg.MouseRow
    c = fg.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    ' make sure the click was on a cell with a button
    If fg.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = fg.Cell(flexcpLeft, r, c) + fg.Cell(flexcpWidth, r, c) - X
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    fg.Cell(flexcpPicture, r, c) = imgBtnDn
    MsgBox "Thanks for clicking my custom button!"
    fg.Cell(flexcpPicture, r, c) = imgBtnUp
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True

End Sub

Private Sub Form_Load()

    ' initialize grid
    fg.Editable = flexEDKbdMouse
    fg.AllowUserResizing = flexResizeBoth
    
    ' add some buttons to the grid
    Dim i%
    For i = 2 To 6
        fg.Cell(flexcpPicture, i, 2) = imgBtnUp
        fg.Cell(flexcpPictureAlignment, i, 2) = flexAlignRightCenter
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - 2 * fg.Left, ScaleHeight - fg.Left - fg.Top
End Sub
