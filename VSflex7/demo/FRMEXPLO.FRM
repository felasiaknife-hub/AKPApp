VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmExplo 
   Caption         =   "FlexExplorer"
   ClientHeight    =   5715
   ClientLeft      =   1995
   ClientTop       =   1410
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VSFlex7LCtl.VSFlexGrid fgList 
      Height          =   4140
      Left            =   3300
      TabIndex        =   8
      Top             =   1500
      Width           =   2940
      _cx             =   5186
      _cy             =   7302
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
   End
   Begin VSFlex7LCtl.VSFlexGrid fgTree 
      Height          =   2640
      Left            =   225
      TabIndex        =   7
      Top             =   2925
      Width           =   2340
      _cx             =   4128
      _cy             =   4657
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
      AutoSearchDelay =   300
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
      WallPaper       =   "frmExplo.frx":0000
      WallPaperAlignment=   0
   End
   Begin VB.DirListBox SubDirs 
      Height          =   540
      Left            =   150
      TabIndex        =   5
      Top             =   1425
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.FileListBox Files 
      Height          =   675
      Left            =   150
      TabIndex        =   4
      Top             =   1950
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.DirListBox Dirs 
      Height          =   540
      Left            =   150
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.DriveListBox Drives 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label picDrag 
      BackColor       =   &H00808080&
      Height          =   2265
      Left            =   2850
      TabIndex        =   6
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lblSelection 
      Caption         =   "Label2"
      Height          =   315
      Left            =   3975
      TabIndex        =   2
      Top             =   825
      Width           =   2415
   End
   Begin VB.Label lblDir 
      Caption         =   "Label1"
      Height          =   540
      Left            =   3150
      TabIndex        =   0
      Top             =   225
      Width           =   3165
   End
   Begin VB.Image imgDrive 
      Height          =   240
      Left            =   450
      Picture         =   "frmExplo.frx":0B5C
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgComputer 
      Height          =   240
      Left            =   150
      Picture         =   "frmExplo.frx":10E6
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFolder 
      Height          =   240
      Left            =   750
      Picture         =   "frmExplo.frx":1670
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFile 
      Height          =   240
      Left            =   1050
      Picture         =   "frmExplo.frx":1BFA
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmExplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fgList_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    ' use shift-clicks to start OLE drag-drop
    If (Shift And 1) = 0 Then Exit Sub
    
    ' tell control to ignore this click, we'll use it ourselves
    Cancel = True
    
    ' start drag-drop (set data object in OLEStartDrag event)
    fgList.OLEDrag
    
End Sub

Private Sub fgList_DragDrop(Source As Control, X As Single, Y As Single)
    
    ' finished resizing: resize fgTree, then fix the other controls
    fgTree.Width = fgTree.Width + fgTree.left + X
    UpdateLayout
    
End Sub

Private Sub fgList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim s$
    If Button = 0 And fgList.Rows > 1 And fgList.MouseRow = 0 Then
        s = "Click headers to sort columns, shift-drag selected files into NotePad or Word."
    End If
    If s <> fgList.ToolTipText Then
        fgList.ToolTipText = s
    End If
End Sub

Private Sub fgList_OLEStartDrag(Data As VSFlex7LCtl.VSDataObject, AllowedEffects As Long)
    
    ' build a list of selected files
    Dim i%, s$, p$
    p = fgTree.RowData(fgTree.Row) & "\"
    For i = 0 To fgList.SelectedRows - 1
        If Len(s) > 0 Then s = s & vbTab
        s = s & p & fgList.TextMatrix(fgList.SelectedRow(i), 0)
    Next
    
    ' put selected files into VSDataObject
    Data.SetData s, vbCFFiles

End Sub

Private Sub fgTree_DragDrop(Source As Control, X As Single, Y As Single)

    ' finished resizing: resize fgTree, then fix the other controls
    fgTree.Width = X
    UpdateLayout

End Sub

Private Sub Form_Load()

    InitControls
    ShowDrives
    
End Sub

Private Sub fgTree_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)

    ' if we're working, ignore this
    If fgTree.Redraw = flexRDNone Then Exit Sub
    
    ' don't allow collapsing/expanding groups of rows (with shift-click)
    If Row < 0 Then
        Cancel = True
        Exit Sub
    End If
    
    ' make sure we're expanding
    If State = flexOutlineCollapsed Then Exit Sub
    
    ' make sure we have to populate branch
    If Row >= fgTree.Rows - 1 Then Exit Sub
    If fgTree.IsSubtotal(Row) = False Then Exit Sub
    If fgTree.TextMatrix(Row + 1, 0) <> "" Then Exit Sub
        
    ' remove dummy branch
    fgTree.RemoveItem Row + 1
    
    ' populate the branch we're about to open
    ExpandRow Row
    
End Sub

Private Sub fgTree_DblClick()

    ' expand/collapse when double-clicking
    If fgTree.IsCollapsed(fgTree.Row) = flexOutlineCollapsed Then
        fgTree.IsCollapsed(fgTree.Row) = flexOutlineExpanded
    Else
        fgTree.IsCollapsed(fgTree.Row) = flexOutlineCollapsed
    End If
    
End Sub

Private Sub fgTree_KeyDown(KeyCode As Integer, Shift As Integer)

    ' use arrow keys to expand/collapse node
    Select Case KeyCode
        Case vbKeyLeft
            fgTree.IsCollapsed(fgTree.Row) = flexOutlineCollapsed
        Case vbKeyRight
            fgTree.IsCollapsed(fgTree.Row) = flexOutlineExpanded
    End Select
        
End Sub

Private Sub Form_Resize()

    UpdateLayout
    
End Sub

Private Sub fgTree_RowColChange()
    
    UpdateFileList
    
End Sub

Private Sub fgList_SelChange()
    
    UpdateFileStats
    
End Sub

Sub ExpandRow(ByVal r&)
    With fgTree
    
        ' in case a floppy is not in the drive...
        On Error Resume Next
        
        ' get path for the given row
        Dim pth$
        pth = .RowData(r)
    
        ' get list of subdirectories
        Dirs.Path = pth & "\"
        If Err <> 0 Then Exit Sub
        
        ' make it fast
        .Redraw = flexRDNone
        MousePointer = 11
    
        ' get information for new items
        Dim iPos&, iLvl&
        iPos = r + 1
        iLvl = .RowOutlineLevel(r) + 1
        
        ' use DirListBox control to list subdirectories
        Dim i%, j%, dr$
        For i = 0 To Dirs.ListCount - 1
    
            ' extract dir from path string
            dr = Dirs.List(i)
            For j = Len(dr) To 1 Step -1
                If Mid(dr, j, 1) = "\" Then
                    dr = Mid(dr, j + 1)
                    Exit For
                End If
            Next
            
            ' add dir
            .AddItem dr, iPos
            .IsSubtotal(iPos) = True
            .RowOutlineLevel(iPos) = iLvl
            .Cell(flexcpPicture, iPos, 0) = imgFolder
        
            ' save path in rowdata
            .RowData(iPos) = Dirs.List(i)
            iPos = iPos + 1
        
            ' if this directory has subdirs, add dummy item
            ' (we'll populate this when it expands)
            SubDirs.Path = Dirs.List(i)
            If SubDirs.ListCount > 0 Then
                .AddItem "", iPos
                .IsSubtotal(iPos) = True
                .RowOutlineLevel(iPos) = iLvl + 1
                iPos = iPos + 1
            End If
        Next
    
        ' collapse the items we just added
        For i = r + 1 To iPos - 1
            If .IsSubtotal(i) Then .IsCollapsed(i) = flexOutlineCollapsed
        Next
    
        ' ready to redraw
        MousePointer = 0
        .Redraw = flexRDBuffered
        .Row = r
    
    End With
End Sub

Sub InitControls()

    ' initialize tree control
    With fgTree
        
        ' structure
        .Cols = 1
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        .left = 50
        
        ' appearance
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered ' << new setting
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarCompleteLeaf
        .Ellipsis = flexEllipsisEnd
        
        ' behavior
        .AllowSelection = False
        .HighLight = flexHighlightWithFocus
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    ' initialize list control
    With fgList
        
        ' structure
        .Cols = 3
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1
        
        ' appearance
        .BorderStyle = flexBorderNone
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .FocusRect = flexFocusNone
        .AllowUserResizing = flexResizeColumns
        .Ellipsis = flexEllipsisEnd
        
        ' behavior
        .ExplorerBar = flexExSortShow
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .AllowUserResizing = flexResizeColumns
        .AutoSearch = flexSearchFromCursor
        
        ' content
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Size (k)"
        .TextMatrix(0, 2) = "Modified"
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColFormat(1) = "#,###.#"
        .ColFormat(2) = "Medium Date"
        .ColDataType(2) = flexDTDate ' << date type needed for sorting
        
    End With
    
    ' initialize label control
    With lblDir
        
        ' appearance
        .FontName = "Tahoma"
        .FontBold = True
        .FontSize = 14
        .BackColor = fgTree.BackColor
        .BorderStyle = 1
        .Caption = ""
        
    End With
    
    ' initialize selection control
    With lblSelection
    
        ' appearance
        .FontName = "Tahoma"
        .FontSize = 8
        .ForeColor = vbBlue
        .Alignment = 1 ' right-aligned
        .BackStyle = 0 ' transparent
        .Caption = ""
    
    End With
    
End Sub

Sub UpdateLayout()
    
    ' make sure this won't crash with small dimensions
    On Error Resume Next
    
    ' initialize parameters
    Dim iBorder%, iTreeWidth%, iLabelHeight%
    iBorder = 2 * Screen.TwipsPerPixelX
    iTreeWidth = fgTree.Width
    iLabelHeight = 600
    
    ' move and position tree control
    With fgTree
        .Move .left, .left, _
              iTreeWidth, ScaleHeight - 2 * .left
    End With
    
    ' move and position main label
    With fgTree
        lblDir.Move 2 * .left + .Width, .top, _
                    ScaleWidth - .Width - 3 * .left, .Height
    End With
    
    ' move and position list
    Dim i%
    i = fgList.Width
    With lblDir
        fgList.Move .left + iBorder, .top + iLabelHeight, _
                    .Width - 2 * iBorder, .Height - iLabelHeight - iBorder
    End With
    fgList.ColWidth(0) = fgList.Width / 2
    fgList.ColWidth(1) = fgList.Width / 4
    
    ' move and position selection label
    With fgList
        lblSelection.Move .left, .top - lblSelection.Height, _
                          .Width, lblSelection.Height
    End With
    
    ' move and position drag control
    With picDrag
        .Move fgList.left - fgTree.left - iBorder, 0, fgTree.left, ScaleHeight
        .MousePointer = flexSizeEW ' to show resizing cursor
        .BackStyle = 0 ' transparent
        .BackColor = BackColor ' don't show this control, only the mouse cursor
    End With
    
End Sub

Sub ShowDrives()
    With fgTree
    
        ' prepare to work
        .Redraw = flexRDNone
        
        ' add computer icon to tree
        .Rows = 0
        .AddItem "My Computer"
        .Cell(flexcpPicture, .Rows - 1) = imgComputer
        .IsSubtotal(.Rows - 1) = True
        
        ' add drives to tree
        Dim i%
        For i = 0 To Drives.ListCount - 1
        
            ' add drive node
            .AddItem Drives.List(i)
            .RowData(.Rows - 1) = left(Drives.List(i), 1) & ":"
            .Cell(flexcpPicture, .Rows - 1) = imgDrive
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 1
            
            ' add hidden node (we'll handle this when the user expands it)
            .AddItem ""
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 2
            
        Next
        
        ' start showing all drives and nothing else
        .Outline 1
        .Row = 0
        fgTree_RowColChange
        
        ' ready
        .Redraw = flexRDBuffered
        
    End With
End Sub

Private Sub UpdateFileStats()

    ' get number of selected files
    Dim iCnt%, i%, sz#
    iCnt = fgList.SelectedRows
    
    ' one or none? easy
    If iCnt < 2 Then
        lblSelection = fgList.Rows - 1 & " Files "
        Exit Sub
    End If
    
    ' calculate size for selection > 1 file using new Aggregate property
    ' row -1 means "all selected rows", col 1 has the file size
    sz = fgList.Aggregate(flexSTSum, -1, 1, -1, 1)
    
    ' the above line is equivalent to this:
    'sz = 0
    'For i = 0 To iCnt - 1
    '    sz = sz + fgList.ValueMatrix(fgList.SelectedRow(i), 1)
    'Next

    ' show label
    lblSelection = iCnt & " Files Selected (" & Format(sz, "#,##0") & "k) "
    
End Sub


Private Sub UpdateFileList()

    ' in case a floppy is not in the drive...
    On Error Resume Next
    
    ' update main label
    lblDir = " " & fgTree
    lblSelection = ""
    
    ' initialize list
    fgList.Rows = 1
    
    ' get selected path
    Dim pth$
    pth = fgTree.RowData(fgTree.Row)
    If Len(pth) = 0 Then Exit Sub
    
    ' save current row
    Dim r%
    r = fgTree.Row
    
    ' use FileList control to get files, quit if the file is not in the drive
    Files.Path = pth & "\"
    If Err <> 0 Then
        lblSelection = "Cannot read " & pth
        Exit Sub
    End If
    
    ' populate the list
    MousePointer = 11
    With fgList
        .Refresh
        .Redraw = flexRDNone
        
        Dim i%, itm$, f$
        For i = 0 To Files.ListCount - 1
            
            ' add this item
            itm = Files.List(i)
            f = pth & "\" & itm
            itm = itm & vbTab & (FileLen(f) / 1024#)
            itm = itm & vbTab & FileDateTime(f)
            .AddItem itm
            
            ' if the user moved, break out
            If .Rows Mod 25 = 0 Then
                .Redraw = flexRDBuffered
                .Refresh
                .Redraw = flexRDNone
                DoEvents
            End If
            If r <> fgTree.Row Then Exit For

        Next
        .Redraw = flexRDDirect
        
    End With
    MousePointer = 0
    
    ' update statistics for this file list
    UpdateFileStats
    
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' start dragging with left button
    If Button <> 1 Then Exit Sub
    
    ' use trick so vertical movement is not visible
    picDrag.Height = 32000
    picDrag.top = -10000
    picDrag.Drag
        
End Sub
