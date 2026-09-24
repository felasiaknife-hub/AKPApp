VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid 7: Date Picker"
   ClientHeight    =   3570
   ClientLeft      =   2760
   ClientTop       =   2715
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6150
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5955
      _cx             =   10504
      _cy             =   5953
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
      Begin MSComCtl2.DTPicker dtPick 
         Height          =   315
         Left            =   3600
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   36473
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtPick_Change()
    
    ' update grid value whenever the data changes
    fg.Text = dtPick.Value
    
End Sub

Private Sub dtPick_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' close date picker when user hits escape or return
    Select Case KeyCode
        Case vbKeyEscape
            fg = dtPick.Tag
            dtPick.Visible = False
        Case vbKeyReturn
            dtPick.Visible = False
    End Select
    
End Sub

Private Sub dtPick_LostFocus()

    ' hide date picker when user is done with it
    dtPick.Visible = False
    
End Sub

Private Sub fg_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    ' don't scroll while editing dates
    If dtPick.Visible Then Cancel = True

End Sub

Private Sub fg_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' don't resize columns while editing dates
    If dtPick.Visible Then Cancel = True

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' if this is a date column, edit it with the date picker control
    If fg.ColDataType(Col) = flexDTDate Then
        
        ' we'll handle the editing ourselves
        Cancel = True
        
        ' position date picker control over cell
        dtPick.Move fg.CellLeft, fg.CellTop, fg.CellWidth, fg.CellHeight
        
        ' initialize value, save original in tag in case user hits escape
        dtPick.Value = fg
        dtPick.Tag = fg
        
        ' show and activate date picker control
        dtPick.Visible = True
        dtPick.SetFocus
        
        ' make it drop down the calendar
        SendKeys "{f4}"
        
    End If

End Sub

Private Sub Form_Load()

    ' initialize grid
    fg.Editable = flexEDKbdMouse
    fg.AllowUserResizing = flexResizeBoth
    fg.Cols = 4
    fg.ColWidth(0) = fg.RowHeight(0)
    fg.ColWidth(1) = 2500
    fg.ColWidth(2) = 1500
    fg.ColWidth(3) = 1500
    fg.TextMatrix(0, 1) = "Name"
    fg.TextMatrix(0, 2) = "Born"
    fg.TextMatrix(0, 3) = "Hired"
    fg.ColDataType(2) = flexDTDate
    fg.ColDataType(3) = flexDTDate
    fg.ColFormat(2) = "Medium Date"
    fg.ColFormat(3) = "Medium Date"
    fg.ColComboList(2) = "Dummy" ' just to show the down-arrow
    fg.ColComboList(3) = "Dummy" ' just to show the down-arrow
    fg.RowHeightMin = 280
    
    ' fill up with dummy data
    Dim r%
    For r = 1 To fg.Rows - 1
        fg.TextMatrix(r, 1) = "Employee " & r
        fg.TextMatrix(r, 2) = Now - 20 * 365 - Rnd * 20 * 365
        fg.TextMatrix(r, 3) = Now - Rnd * 10 * 365
    Next
    
End Sub

Private Sub Form_Resize()

    ' resize grid to fill form
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - 2 * fg.Left, ScaleHeight - fg.Top - fg.Left

End Sub
