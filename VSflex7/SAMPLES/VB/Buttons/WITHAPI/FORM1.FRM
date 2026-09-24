VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   6585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6585
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2475
      _cx             =   4366
      _cy             =   3942
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
   Begin VB.Image imgBtnBig 
      Height          =   480
      Left            =   4860
      Picture         =   "Form1.frx":011A
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBtn 
      Height          =   180
      Left            =   4500
      Picture         =   "Form1.frx":0424
      Top             =   135
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXHSCROLL = 21
Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10


Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' if the cell has a button, show it
    If fg.Cell(flexcpData, Row, Col) = "Btn" Then
        fg.ComboList = "..."
    Else
        fg.ComboList = ""
    End If
    
End Sub

Private Sub fg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    ' process click on custom button
    Form2.Show 1
    
End Sub

Private Sub fg_DrawCell(ByVal hdc As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    
    ' if the cell has no button, don't do anything
    If fg.Cell(flexcpData, Row, Col) <> "Btn" Then Exit Sub

    ' if the cell has the focus, no need to draw it
    If Row = fg.Row And Col = fg.Col And fg.hWnd = GetFocus() Then Exit Sub
    
    ' draw the custom button frame
    Dim rc As RECT
    rc.Left = Right - GetSystemMetrics(SM_CXHSCROLL) - 2
    rc.Top = Top
    rc.Right = Right - 1
    rc.Bottom = Bottom - 1
    DrawFrameControl hdc, rc, DFC_BUTTON, DFCS_BUTTONPUSH
    
    ' draw the custom button image
    DrawIcon hdc, rc.Left + 3, (rc.Top + rc.Bottom) / 2 - 5, imgBtnBig.Picture.Handle
    
End Sub

Private Sub Form_Load()

    ' intialize grid
    fg.Editable = flexEDKbdMouse
    fg.AllowUserResizing = flexResizeBoth
    fg.OwnerDraw = flexODOver
    fg.CellButtonPicture = imgBtn
    
    ' spread a few buttons around
    Dim i%
    For i = 2 To 6
        fg.Cell(flexcpData, i, 1) = "Btn"
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - 2 * fg.Left, ScaleHeight - fg.Left - fg.Top
End Sub

