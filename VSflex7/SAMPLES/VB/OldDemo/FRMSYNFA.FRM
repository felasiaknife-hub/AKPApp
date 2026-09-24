VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6.OCX"
Begin VB.Form frmSynFa 
   Caption         =   "Synchronized To a Flex"
   ClientHeight    =   4860
   ClientLeft      =   2505
   ClientTop       =   2040
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   6315
   Begin VB.TextBox lblHelp 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   36
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   36
      Width           =   6132
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   3615
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   750
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   6376
      _ConvInfo       =   -1
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   3
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   -1  'True
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   3600
      Index           =   1
      Left            =   4860
      TabIndex        =   1
      Top             =   720
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   6350
      _ConvInfo       =   -1
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   0
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   -1  'True
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   315
      Index           =   2
      Left            =   75
      TabIndex        =   2
      Top             =   4440
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   556
      _ConvInfo       =   -1
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   0
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   -1  'True
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   315
      Index           =   3
      Left            =   4905
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _ConvInfo       =   -1
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   0
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   0   'False
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   -1  'True
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
End
Attribute VB_Name = "frmSynFa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UpdateTotals()
    Dim r&, c&, tot!
    
    For c = 1 To fa(0).Cols - 2
        tot = 0
        For r = 1 To fa(0).Rows - 2
            tot = tot + fa(0).ValueMatrix(r, c)
        Next
        fa(0).TextMatrix(r, c) = tot
    Next

    For r = 1 To fa(0).Rows - 1
        tot = 0
        For c = 1 To fa(0).Cols - 2
            tot = tot + fa(0).ValueMatrix(r, c)
        Next
        fa(0).TextMatrix(r, c) = tot
    Next
End Sub

Private Sub fa_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i%
    
    ' update totals
    UpdateTotals
    
    ' refresh other views after edits
    For i = 0 To 3
        If i <> Index Then fa(i).Refresh
    Next
End Sub

Private Sub fa_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = fa(0).Rows - 1 Or Col = fa(0).Cols - 1 Then Cancel = True
End Sub

Private Sub fa_GotFocus(Index As Integer)
    If Index <> 0 Then fa(0).SetFocus
End Sub


Private Sub fa_Scroll(Index As Integer)
    
    ' synchronize scrolling
    If Index = 0 Then
        fa(1).TopRow = fa(0).TopRow
        fa(2).LeftCol = fa(0).LeftCol
    End If
End Sub



Private Sub Form_Load()
    
    fa(0).Cell(flexcpText, fa(0).Rows - 1, 0) = "Total"
    fa(0).Cell(flexcpText, 0, fa(0).Cols - 1) = "Total"
    UpdateTotals
    
    fa(1).BindToArray fa(0)
    fa(2).BindToArray fa(0)
    fa(3).BindToArray fa(0)
    
    fa(1).ColWidth(0) = 0
    fa(2).RowHeight(0) = 0
    fa(3).ColWidth(0) = 0
    fa(3).RowHeight(0) = 0
    
    fa(1).Width = fa(1).ColWidth(1) + 50
    fa(2).Height = fa(2).RowHeight(1) + 50
    fa(3).Width = fa(1).Width
    fa(3).Height = fa(3).Height

End Sub


Private Sub Form_Resize()
    Dim w!, h!, l!, t!
    l = fa(0).Left
    
    ' adjust label position
    w = ScaleWidth - 2 * l
    lblHelp.Top = l
    If w > 0 Then lblHelp.Width = w
    
    ' adjust grid positions
    t = 2 * l + lblHelp.Height
    w = ScaleWidth - fa(1).Width - 3 * l
    h = ScaleHeight - fa(2).Height - t - 2 * l
    If (w > 0) And (h > 0) Then fa(0).Move l, t, w, h
    fa(1).Move ScaleWidth - fa(1).Width - l, t, fa(1).Width, h
    fa(2).Move l, fa(0).Top + fa(0).Height + l, w
    fa(3).Move fa(1).Left, fa(2).Top, fa(1).Width, fa(2).Height

    ' make sure we're showing the right stuff
    fa_Scroll 0
    fa(1).LeftCol = fa(1).Cols - 1
    fa(2).TopRow = fa(2).Rows - 1
    fa(3).LeftCol = fa(3).Cols - 1
    fa(3).TopRow = fa(3).Rows - 1
End Sub


