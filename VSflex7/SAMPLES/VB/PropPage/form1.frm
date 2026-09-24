VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Property Page"
   ClientHeight    =   4935
   ClientLeft      =   2670
   ClientTop       =   2070
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4320
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   1290
      Left            =   0
      TabIndex        =   3
      Top             =   825
      Width           =   3990
      _cx             =   7038
      _cy             =   2275
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
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   75
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frm 
      Caption         =   "Display Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      Begin VB.OptionButton optAlpha 
         Caption         =   "&Alphabetic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton optCat 
         Caption         =   "&Categorized"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1875
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1440
      End
   End
   Begin VB.Image imgFontPick 
      Height          =   240
      Left            =   450
      Picture         =   "form1.frx":0000
      Top             =   2850
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgColorPick 
      Height          =   240
      Left            =   75
      Picture         =   "form1.frx":014A
      Top             =   2850
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' property type
Private Type Property_Type
    category As String
    name As String
    ptype As Integer
    value As Variant
End Type

Private Enum PropertyType_Type
    ptNil
    ptcolor
    ptBool
    ptFont
    ptFontName
    ptValue
End Enum

' property vector
Dim g_Props(30) As Property_Type

' string to hold font list
Dim g_FontList$

' API Declarations (for OwnerDraw cells)
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const ETO_OPAQUE = 2
Private Const BLACK_BRUSH = 4

Private Sub DisplayCategorized()
    
    ' freeze to avoid flicker
    fg.Redraw = flexRDNone
    
    ' remove any existing subtotals (groups)
    fg.Subtotal flexSTClear
    
    ' sort by category, then by property name
    fg.Select 1, 1, 1, 2
    fg.Sort = flexSortStringAscending
    
    ' add subtotals (groups) by category (col 1)
    fg.Subtotal flexSTNone, 1, , , fg.GridColor, , True
    
    ' show outline column
    fg.ColHidden(0) = False
    
    ' to look nice
    fg.GridLines = flexGridFlatVert
    
    ' reset display
    fg.TopRow = 1
    fg.Select 2, fg.Cols - 1
    fg.Redraw = flexRDBuffered

End Sub

Private Sub DisplayAlphabetic()
    
    ' freeze to avoid flicker
    fg.Redraw = flexRDNone
    
    ' remove any existing subtotals (groups)
    fg.Subtotal flexSTClear
    
    ' sort by property name
    fg.Col = 2
    fg.Sort = flexSortStringAscending
    
    ' hide outline column
    fg.ColHidden(0) = True
    
    ' to look nice
    fg.GridLines = flexGridFlat
    
    ' reset display
    fg.TopRow = 1
    fg.Select 1, fg.Cols - 1
    fg.Redraw = flexRDBuffered

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' we can't edit total rows or label columns
    If fg.IsSubtotal(Row) Or Col <> fg.Cols - 1 Then
        Cancel = True
        Exit Sub
    End If
    
    ' assume regular editing
    fg.ComboList = ""
    
    ' setup to edit based on property type
    Select Case g_Props(fg.RowData(Row)).ptype
    
        ' font name gets a combo
        Case ptFontName
            fg.ComboList = g_FontList

        ' fonts get a pick button
        Case ptFont
            fg.ComboList = "..."
            fg.CellButtonPicture = imgFontPick
    
        ' colors get a different pick button
        Case ptcolor
            fg.ComboList = "..."
            fg.CellButtonPicture = imgColorPick
    
        ' booleans get a pick list
        Case ptBool
            fg.ComboList = "True|False"
    End Select

    ' use automatic double-click for editing text, manual for lists
    If Len(fg.ComboList) Then
        fg.Editable = flexEDKbd
    Else
        fg.Editable = flexEDKbdMouse
    End If

End Sub

Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    ' user can select only the last column
    With fg
        If .Redraw <> flexRDNone And NewCol <> .Cols - 1 Then
            Cancel = True
            .Select NewRow, .Cols - 1
        End If
    End With
    
End Sub

Private Sub fg_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' don't resize outline column
    If Col = 0 Then Cancel = True
    
End Sub

Private Sub fg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    cmDlg.CancelError = False
    
    ' clicked the button to edit a color?
    If g_Props(fg.RowData(Row)).ptype = ptcolor Then
    
        ' position palette form below current cell
        frmPalette.Move Left + fg.Left + fg.ColPos(Col) + (Width - ScaleWidth), _
                        Top + fg.Top + fg.RowPos(Row) + fg.RowHeight(Row) + (Height - ScaleHeight)

                ' show the palette
                ' we show the palette as a modeless dialog and wait until the user is done with it, either
                ' by clicking a color, pressing ESC, or just by activating some other window.
                ' alternatively, we could show it modally and then we could remove the While statement.
                '
        frmPalette.Show
        While frmPalette.Visible
            DoEvents
        Wend
        
        ' if the user picked a value, it's in the Tag property
        If frmPalette.Tag <> "" Then fg.TextMatrix(Row, fg.Cols - 1) = FormatColor(frmPalette.Tag)
        Unload frmPalette
        
        ' we could use the common color dialog instead, but it has no support for system colors
        'cmDlg.Flags = cdlCCRGBInit
        'cmDlg.Color = Val(fg.TextMatrix(Row, fg.Cols - 1))
        'cmDlg.ShowColor
        'fg.TextMatrix(Row, fg.Cols - 1) = FormatColor(cmDlg.Color)
        
    ' clicked the button to edit a font?
    ElseIf g_Props(fg.RowData(Row)).ptype = ptFont Then
        cmDlg.Flags = cdlCFBoth Or cdlCCRGBInit Or cdlCFEffects
        cmDlg.FontName = fg.Cell(flexcpFontName, Row, fg.Cols - 1)
        cmDlg.FontBold = fg.Cell(flexcpFontBold, Row, fg.Cols - 1)
        cmDlg.FontItalic = fg.Cell(flexcpFontItalic, Row, fg.Cols - 1)
        cmDlg.FontSize = fg.Cell(flexcpFontSize, Row, fg.Cols - 1)
        cmDlg.FontUnderline = fg.Cell(flexcpFontUnderline, Row, fg.Cols - 1)
        cmDlg.FontStrikethru = fg.Cell(flexcpFontStrikethru, Row, fg.Cols - 1)
        cmDlg.Color = fg.Cell(flexcpForeColor, Row, fg.Cols - 1)
        cmDlg.ShowFont
        fg.TextMatrix(Row, fg.Cols - 1) = cmDlg.FontName
        
        ' format cell to show font
        fg.Cell(flexcpFontName, Row, fg.Cols - 1) = cmDlg.FontName
        fg.Cell(flexcpFontBold, Row, fg.Cols - 1) = cmDlg.FontBold
        fg.Cell(flexcpFontItalic, Row, fg.Cols - 1) = cmDlg.FontItalic
        fg.Cell(flexcpFontSize, Row, fg.Cols - 1) = cmDlg.FontSize
        fg.Cell(flexcpFontUnderline, Row, fg.Cols - 1) = cmDlg.FontUnderline
        fg.Cell(flexcpFontStrikethru, Row, fg.Cols - 1) = cmDlg.FontStrikethru
        fg.Cell(flexcpForeColor, Row, fg.Cols - 1) = cmDlg.Color
    
    End If
    
End Sub

Private Sub fg_DblClick()

    ' double-clicking on a group collapses/expands it
    Dim r%
    r = fg.MouseRow
    If fg.IsSubtotal(r) Then
        If fg.IsCollapsed(r) = flexOutlineCollapsed Then
            fg.IsCollapsed(r) = flexOutlineExpanded
        Else
            fg.IsCollapsed(r) = flexOutlineCollapsed
        End If
        fg.Tag = ""
        
    ' double-clicking on regular cells edits them
    Else
        fg.Tag = "*"
        fg.EditCell
    End If
    
End Sub

Private Sub fg_DrawCell(ByVal hdc As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)

    ' only need to custom draw color selection cells
    If Col <> 3 Then Exit Sub
    If g_Props(fg.RowData(Row)).ptype <> ptcolor Then Exit Sub
    
    ' build color rectangle
    Dim rc As RECT
    rc.Left = Left + 2
    rc.Right = Left + 15
    rc.Top = Top + 2
    rc.Bottom = Bottom - 3
    
    ' translate color
    Dim clr&
    clr = Val(fg.TextMatrix(Row, fg.Cols - 1))
    If (clr And &H80000000) Then
        clr = GetSysColor(clr And &HFF)
    End If
    
    ' paint rectangle
    clr = SetBkColor(hdc, clr)
    ExtTextOut hdc, 0, 0, ETO_OPAQUE, rc, 0, 0, 0
    SetBkColor hdc, clr
    
    ' frame rectangle
    FrameRect hdc, rc, GetStockObject(BLACK_BRUSH)
    
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r%
    
    ' special handling for cursor keys
    Select Case KeyCode

        ' collapse/expand with cursor keys
        Case vbKeyLeft, vbKeyHome
            If fg.IsSubtotal(fg.Row) Then fg.IsCollapsed(fg.Row) = flexOutlineCollapsed
            If fg.Col <> fg.Cols - 1 Then fg.Col = fg.Cols - 1
            KeyCode = 0
        Case vbKeyRight, vbKeyEnd
            If fg.IsSubtotal(fg.Row) Then fg.IsCollapsed(fg.Row) = flexOutlineExpanded
            If fg.Col <> fg.Cols - 1 Then fg.Col = fg.Cols - 1
            KeyCode = 0
                        
        ' when pushing control+ASCII, look for property
        Case Else
            If Shift >= 2 And KeyCode >= Asc("A") And KeyCode <= Asc("Z") Then
            
                ' look from current row down to bottom
                For r = fg.Row + 1 To fg.Rows - 1
                    If Not fg.RowHidden(r) Then
                        If UCase(Left(fg.TextMatrix(r, fg.Cols - 2), 1)) = Chr(KeyCode) Then
                            fg.Select r, fg.Cols - 1
                            fg.ShowCell r, fg.Cols - 1
                            KeyCode = 0
                            Exit For
                        End If
                    End If
                Next
                
                ' not found, so look from top down to current - 1
                For r = fg.FixedRows To fg.Row - 1
                    If Not fg.RowHidden(r) Then
                        If UCase(Left(fg.TextMatrix(r, fg.Cols - 2), 1)) = Chr(KeyCode) Then
                            fg.Select r, fg.Cols - 1
                            fg.ShowCell r, fg.Cols - 1
                            KeyCode = 0
                            Exit For
                        End If
                    End If
                Next
            End If
    End Select
    
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' if this is a list, double-clicking selects the next item
    If Len(fg.Tag) > 0 And Len(fg.ComboList) > 0 And fg.ComboList <> "..." Then
        If ComboNext(Row, Col) Then Cancel = True
    End If
    fg.Tag = ""
    
End Sub

Private Function ComboNext(Row&, Col&) As Boolean

    ' get current list, trim combo pipe if any
    Dim s$, i%
    s = fg.ComboList
    If Left(s, 1) = "|" Then s = Mid(s, 2)
    
    ' look for current text in list, fail if not found
    i = InStr(s, fg.TextMatrix(Row, Col))
    If i <= 0 Then Exit Function
    
    ' look for next choice
    i = InStr(i, s, "|")
    If (i > 0) Then s = Mid(s, i + 1)
    
    ' trim excess
    i = InStr(s, "|")
    If i > 0 Then s = Left(s, i - 1)
    
    ' set new entry
    fg.TextMatrix(Row, Col) = s
    ComboNext = True
    
End Function

Private Sub Form_Load()

    ' build font list to display in combo
    Dim i%
    For i = 0 To Screen.FontCount - 1
        g_FontList = g_FontList & "|" & Screen.Fonts(i)
    Next

    ' initialize property list
    InitPropertyList
    
    ' initialize control
    fg.Rows = 1                                 ' start empty
    fg.Cols = 4                                 ' outline, category, property, value
    fg.TextMatrix(0, 2) = "Property"            ' column titles
    fg.TextMatrix(0, 3) = "Value"               ' column titles
    fg.Editable = flexEDKbd                     ' double-clicks start editing
    fg.OwnerDraw = flexODOver                   ' use ownerdraw to show colors
    fg.OutlineCol = 0                           ' set outline column properties
    fg.ColWidth(0) = 230                        ' narrow outline column
    fg.OutlineBar = flexOutlineBarSymbolsLeaf   ' no tree, just symbols
    fg.ColHidden(1) = True                      ' hide categories
    fg.MergeCells = flexMergeSpill              ' allow categories to spill into property column
    fg.ColAlignment(-1) = flexAlignLeftTop      ' align all to left
    fg.AllowSelection = False                   ' select a single cell at a time
    fg.AllowUserResizing = flexResizeColumns    ' give user freedom
    fg.ScrollTrack = True                       ' scroll as the user drags the scroll thumb
    fg.FixedCols = 0                            ' to look nice
    fg.ExtendLastCol = True
    fg.Ellipsis = flexEllipsisEnd
    fg.BackColorBkg = fg.GridColor
    fg.HighLight = flexHighlightNever
    
    ' populate control
    For i = 1 To UBound(g_Props)
        fg.AddItem vbTab & g_Props(i).category & vbTab & g_Props(i).name & vbTab & g_Props(i).value
        fg.RowData(i) = i ' keep property index because we will be sorting this
    Next

    ' do an autosize on property names
    fg.AutoSize fg.Cols - 2, , , 300
    
    ' initialize display
    DisplayCategorized
    
End Sub

Private Sub Form_Resize()
    With fg
        .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight - .Top - .Left
        frm.Move .Left, frm.Top, .Width
    End With
End Sub

Private Sub optAlpha_Click()
    DisplayAlphabetic
End Sub

Private Sub optCat_Click()
    DisplayCategorized
End Sub

Private Function FormatColor(clr&) As String
    
    ' translate value into fixed-length hex
    Dim s$
    s = Hex(clr)
    If Len(s) < 8 Then s = String(8 - Len(s), "0") & s

    ' prepend 'H' and some spaces to fit owner-drawn color box
    FormatColor = "     &H" & s & "&"

End Function

Private Sub InitPropertyList()
    Dim i%
    
    For i = 1 To UBound(g_Props)
        
        ' randomize category
        Select Case (Rnd * 1000) Mod 4
            Case 0
                g_Props(i).category = "Appearance"
            Case 1
                g_Props(i).category = "Behavior"
            Case 2
                g_Props(i).category = "Structure"
            Case Else
                g_Props(i).category = "Miscellaneous"
        End Select
        
        ' randomize name/value
        Select Case (Rnd * 1000) Mod 5
            Case 0
                g_Props(i).ptype = ptcolor
                g_Props(i).name = "Color_" & Format(i, "00")
                g_Props(i).value = FormatColor(RGB(Rnd * 255, Rnd * 255, Rnd * 255))
            Case 1
                g_Props(i).ptype = ptBool
                g_Props(i).name = "Bool_" & Format(i, "00")
                g_Props(i).value = True
            Case 2
                g_Props(i).ptype = ptFont
                g_Props(i).name = "Font_" & Format(i, "00")
                Set g_Props(i).value = Font
            Case 3
                g_Props(i).ptype = ptFontName
                g_Props(i).name = "FontName_" & Format(i, "00")
                g_Props(i).value = Font.name
            Case Else
                g_Props(i).ptype = ptValue
                g_Props(i).name = "Value_" & Format(i, "00")
                g_Props(i).value = 123.123
        End Select
    Next
End Sub

