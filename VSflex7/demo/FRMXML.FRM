VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmXML 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VSFlexGrid Pro 7.0: Parsing XML Documents"
   ClientHeight    =   6585
   ClientLeft      =   150
   ClientTop       =   1440
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse XML"
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtXML 
      Enabled         =   0   'False
      Height          =   4215
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmXML.frx":0000
      Top             =   2040
      Width           =   4575
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4335
      Left            =   6420
      TabIndex        =   0
      Top             =   1980
      Width           =   5055
      _cx             =   8916
      _cy             =   7646
      _ConvInfo       =   1
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Content of XML file: Test.XML"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1740
      Width           =   3195
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   7560
      X2              =   420
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   300
      Shape           =   3  'Circle
      Top             =   1140
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   420
      Picture         =   "frmXML.frx":03EE
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdParse_Click()
       
    MousePointer = 11
    fg.Rows = 1
    
    ' read XML document into the grid
    Dim xml As New MSXML.XMLDocument
    xml.URL = App.Path & "\test.xml"
    DumpXMLNode xml.root, 0
    
    fg.AutoSize 0, 1, , 200
    
    Set xml = Nothing
    
    MousePointer = 0
    
End Sub


Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
    If fg.IsSubtotal(fg.Row) Then
        Select Case KeyCode
            Case vbKeyLeft
                fg.IsCollapsed(fg.Row) = flexOutlineCollapsed
            Case vbKeyRight
                fg.IsCollapsed(fg.Row) = flexOutlineExpanded
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
    

End Sub

Private Sub DumpXMLNode(node As IXMLElement2, lvl%)
    Dim iChildren%, i%
    
    If Not (node.Children Is Nothing) Then
        iChildren = node.Children.length
    End If
    
    If iChildren = 1 Then
        fg.AddItem node.tagName & vbTab & node.text
        fg.IsSubtotal(fg.Rows - 1) = True
        fg.RowOutlineLevel(fg.Rows - 1) = lvl
        fg.Cell(flexcpBackColor, fg.Rows - 1, 1) = &HC0E0FF   '&HF0F0F0
    Else
        fg.AddItem node.tagName
        fg.IsSubtotal(fg.Rows - 1) = True
        fg.RowOutlineLevel(fg.Rows - 1) = lvl
        For i = 0 To iChildren - 1
            DumpXMLNode node.Children.Item(i), lvl + 1
        Next
    End If
    
End Sub


