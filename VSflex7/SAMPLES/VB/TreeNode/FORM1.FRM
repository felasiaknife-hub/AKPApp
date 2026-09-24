VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid 7: Tree-Handling with the VSFlexNode Object"
   ClientHeight    =   5610
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7020
   Begin VB.CheckBox chkDrag 
      Caption         =   "&Dragging"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   540
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "ñ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4860
      TabIndex        =   14
      ToolTipText     =   "Promote"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "ò"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      ToolTipText     =   "Demote"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   3840
      TabIndex        =   11
      ToolTipText     =   "Move Last"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   3540
      TabIndex        =   10
      ToolTipText     =   "Move First"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "â"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   3180
      TabIndex        =   9
      ToolTipText     =   "Move Down"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "á"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Move Up"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "à"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   "Demote"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "ß"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2220
      TabIndex        =   6
      ToolTipText     =   "Promote"
      Top             =   480
      Width           =   315
   End
   Begin VB.ComboBox cmbAdd 
      Height          =   315
      ItemData        =   "form1.frx":0000
      Left            =   4860
      List            =   "form1.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   60
      Width           =   1500
   End
   Begin VB.ComboBox cmbSelect 
      Height          =   315
      ItemData        =   "form1.frx":006A
      Left            =   2220
      List            =   "form1.frx":0086
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1500
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3615
      Left            =   180
      TabIndex        =   0
      Top             =   900
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sort"
      Height          =   195
      Left            =   4500
      TabIndex        =   15
      Top             =   540
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Move"
      Height          =   195
      Left            =   1680
      TabIndex        =   12
      Top             =   540
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Add"
      Height          =   195
      Left            =   4500
      TabIndex        =   5
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' to handle node dragging
Const DRAGTOL = 100         ' mouse movement before dragging starts
Private Type DRAGINFO
    bDragging As Boolean    ' currently dragging
    bCheckDrag As Boolean   ' currently checking mouse to start dragging
    lSrc As Long            ' row being dragged
    xDown As Long           ' mouse down position
    yDown As Long           ' mouse down position
End Type
Dim g_DragInfo As DRAGINFO

Private Sub cmdRefresh_Click()

    ' read original XML document into grid
    ReadXML
    
End Sub

Private Sub cmbSelect_Click()

    ' get current row's node
    Dim nd As VSFlexNode
    Set nd = fg.GetNode

    ' get relative selected by the user
    Set nd = nd.GetNode(cmbSelect.ListIndex)
    
    ' if failed, print message
    If nd Is Nothing Then
        Debug.Print "Invalid request for this node"
        Exit Sub
    End If
    
    ' select and show relative
    nd.Select
    nd.EnsureVisible
    
End Sub

Private Sub cmbAdd_Click()
    
    ' get current row's node
    Dim nd As VSFlexNode
    Set nd = fg.GetNode
    
    ' add relative as requested by user
    ' (could be a child or a sibling)
    If cmbAdd.ListIndex >= 0 Then
        nd.AddNode cmbAdd.ListIndex + 2, cmbAdd.Text
    End If
    
End Sub

Private Sub cmdMove_Click(Index As Integer)
    
    ' get current row's node
    Dim nd As VSFlexNode
    Set nd = fg.GetNode
    
    ' apply movement selected by the user
    ' (this will move the selected node)
    Select Case Index
        Case 0: nd.Move flexNMOut
        Case 1: nd.Move flexNMIn
        Case 2: nd.Move flexNMUp
        Case 3: nd.Move flexNMDown
        Case 4: nd.Move flexNMFirst
        Case 5: nd.Move flexNMLast
    End Select
    
    ' make sure the node is still visible
    nd.EnsureVisible
    
End Sub

Private Sub cmdSort_Click(Index As Integer)
    
    ' get current row's node
    Dim nd As VSFlexNode
    Set nd = fg.GetNode
    
    ' apply sorting selected by the user
    ' (this will sort the selected node's children)
    Select Case Index
        Case 0: nd.Sort flexSortGenericAscending
        Case 1: nd.Sort flexSortGenericDescending
    End Select
    
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)

    ' use keypboard to expand/collapse nodes and to cancel dragging
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight
            
            ' do we have a node with children?
            Dim nd As VSFlexNode
            Set nd = fg.GetNode
            If nd Is Nothing Then Exit Sub
            If nd.Children = 0 Then Exit Sub
            
            ' we do, so expand or collapse it
            nd.Expanded = (KeyCode = vbKeyRight)
            KeyCode = 0
                
        ' cancel dragging
        Case vbKeyEscape
            If g_DragInfo.bDragging Then
                g_DragInfo.bDragging = False
                fg.MousePointer = flexDefault
                fg.Cell(flexcpCustomFormat, g_DragInfo.lSrc, 0) = False
                KeyCode = 0
            End If
    End Select
    
End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' left button, no shift: start tracking mouse to drag
    If (Button = 1) And (Shift = 0) Then
        If (chkDrag.Value <> 0) And (g_DragInfo.bDragging = False) Then
            
            ' save current row and mouse position
            g_DragInfo.lSrc = fg.Row
            g_DragInfo.xDown = X
            g_DragInfo.yDown = Y
            
            ' start checking
            g_DragInfo.bCheckDrag = True
        End If
    End If
End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' if checking and the user moved past our tolerance, start dragging
    If g_DragInfo.bCheckDrag And Button Then
        If Abs(X - g_DragInfo.xDown) + Abs(Y - g_DragInfo.yDown) > DRAGTOL Then
        
            ' update flags
            g_DragInfo.bDragging = True
            g_DragInfo.bCheckDrag = False
            
            ' set cursor and highlight node
            fg.MousePointer = flexPointerMove
            fg.Cell(flexcpBackColor, g_DragInfo.lSrc, 0) = vbYellow
        End If
    End If
End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' if we were dragging, then do it
    If g_DragInfo.bDragging Then
    
        ' stop dragging
        fg.Cell(flexcpCustomFormat, g_DragInfo.lSrc, 0) = False
        g_DragInfo.bDragging = False
        fg.MousePointer = flexDefault
        
        ' move node into new parent node
        Dim nd As VSFlexNode
        Set nd = fg.GetNode(g_DragInfo.lSrc)
        nd.Move flexNMChildOf, fg.GetNode
        nd.Select

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
        .Editable = flexEDKbdMouse
                
        ' outline
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarSimpleLeaf
        .MergeCells = flexMergeOutline
        
        ' other
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .GridLines = flexGridFlatVert
    End With
    
    ' read XML document into the grid
    ReadXML
    
End Sub

Private Sub Form_Resize()

    ' resize grid to fit form
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - fg.Left * 2, ScaleHeight - fg.Top - fg.Left
    
End Sub

Private Sub ReadXMLNode(node As IXMLElement2, lvl%)
    
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
        ReadXMLNode node.Children.Item(i), lvl + 1
    Next
    
End Sub

Private Sub ReadXML()
    Dim xml As New MSXML.XMLDocument
    xml.url = App.Path & "\test.xml"
    fg.Rows = 1
    fg.Redraw = flexRDNone
    ReadXMLNode xml.root, 0
    fg.AutoSize 0, 1, , 200
    fg.Redraw = flexRDBuffered
End Sub
