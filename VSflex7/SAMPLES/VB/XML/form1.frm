VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid 7: Parsing XML Documents"
   ClientHeight    =   4125
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6705
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5955
      _cx             =   10504
      _cy             =   6376
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
    If fg.IsSubtotal(fg.Row) Then
        Select Case KeyCode
            Case vbKeyLeft
                fg.IsCollapsed(fg.Row) = flexOutlineCollapsed
                KeyCode = 0
            Case vbKeyRight
                fg.IsCollapsed(fg.Row) = flexOutlineExpanded
                KeyCode = 0
        End Select
    End If
End Sub

Private Sub Form_Load()
    
    ' initialize grid
    With fg
    
        ' layout
        .Rows = 1
        .Cols = 2
        .FixedCols = 0
        .ExtendLastCol = True
        .TextMatrix(0, 0) = "Tag"
        .TextMatrix(0, 1) = "Value"
        .ColAlignment(-1) = flexAlignLeftTop
        .ScrollBars = flexScrollBarVertical
                
        ' outline
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarSimpleLeaf
        .MergeCells = flexMergeOutline
        
        ' other
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .HighLight = flexHighlightNever
        .GridLines = flexGridFlatVert
    End With
    
    ' read XML document into the grid
    On Error Resume Next
    Dim xml As New MSXML.XMLDocument
    If Err <> 0 Then
        MsgBox "Cannot create XML parser. Please make sure MSXML.DLL is registered.", vbInformation
    Else
        xml.url = App.Path & "\test.xml"
        If Err <> 0 Then
            MsgBox "Cannot open test file 'test.xml'.", vbInformation
        Else
            fg.Redraw = flexRDNone
            DumpXMLNode xml.root, 0
            fg.Redraw = flexRDBuffered
            fg.AutoSize 0, 1, , 200
            If Err <> 0 Then MsgBox "Syntax errors were found in the XML file.", vbInformation
        End If
    End If
End Sub

Private Sub DumpXMLNode(node As IXMLElement2, lvl%)
    
    ' ignore empty elements
    If node.Children Is Nothing Then Exit Sub
    
    ' get child count
    Dim iChildren%, i%
    iChildren = node.Children.length
    
    ' only nodes of type zero have child nodes
    If node.Children.Item(0).Type <> 0 Then iChildren = 0

    ' no children? add leaf node
    If iChildren = 0 Then
        fg.AddItem node.tagName & vbTab & node.Text
        fg.Cell(flexcpBackColor, fg.Rows - 1, 1) = &HF0F0F0
            
    ' have children? add branch node
    Else
        fg.AddItem node.tagName
    End If
    
    ' either way, they are nodes
    fg.IsSubtotal(fg.Rows - 1) = True
    fg.RowOutlineLevel(fg.Rows - 1) = lvl
    
    ' dump child nodes
    For i = 0 To iChildren - 1
        DumpXMLNode node.Children.Item(i), lvl + 1
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - fg.Left * 2, ScaleHeight - fg.Top - fg.Left
End Sub

