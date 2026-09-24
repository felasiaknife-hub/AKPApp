VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmTotals 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VSFlexGrid Pro 6.0 - SubTotals Demo"
   ClientHeight    =   7110
   ClientLeft      =   630
   ClientTop       =   1170
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7110
   ScaleWidth      =   10635
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Automatic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1050
      TabIndex        =   0
      Top             =   990
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Flip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2145
      TabIndex        =   1
      Top             =   960
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2145
      TabIndex        =   2
      Text            =   "0"
      Top             =   1350
      Width           =   1200
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   4995
      Left            =   300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   10155
      _cx             =   5080
      _cy             =   5080
      _ConvInfo       =   -1
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
      ForeColor       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   16711680
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   89
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "|^Region           |^Product             |^Employee        |>Sales                  "
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
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
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Explorer Look"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MergeCells"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   4
      Top             =   1320
      Width           =   1470
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   405
      Shape           =   3  'Circle
      Top             =   780
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   525
      X2              =   4560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image imgConvert 
      Height          =   270
      Index           =   8
      Left            =   765
      Picture         =   "frmTotal.frx":0000
      Top             =   480
      Width           =   1395
   End
   Begin VB.Image imgConvert 
      Height          =   300
      Index           =   1
      Left            =   5640
      Picture         =   "frmTotal.frx":0E76
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Totals using"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1245
      TabIndex        =   6
      Top             =   1380
      Width           =   1050
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   5520
      X2              =   10320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   780
      Width           =   135
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Fmt As String

Sub SortAndShow()
    Dim r%, c%
    
    fa.Redraw = False
    ' save cursor
    r = fa.Row
    c = fa.Col
    
    ' sort
    fa.Col = 0
    fa.ColSel = fa.Cols - 1
    fa.Sort = 1

    ' remove old totals, calculate new totals
    fa.Subtotal flexSTClear
    
    If Check3 Then
     fa.Subtotal Combo1.ItemData(Combo1.ListIndex), -1, 4, m_Fmt, , RGB(255, 0, 0), True
     fa.Subtotal Combo1.ItemData(Combo1.ListIndex), 1, 4, m_Fmt, , RGB(0, 0, 255), True
     fa.Subtotal Combo1.ItemData(Combo1.ListIndex), 2, 4, m_Fmt, , RGB(0, 0, 255), True
     fa.Subtotal Combo1.ItemData(Combo1.ListIndex), 3, 4, m_Fmt, , RGB(0, 0, 255), True
    End If
    ' restore cursor
    fa.Row = r
    fa.Col = c
    fa.Redraw = True
    
End Sub


Sub TotalInsert(tot#, colg%, i%, j%, clr&, fmt$)
    Dim s$
    Dim k%
    
    ' built total item
    If colg < 0 Then
        s = "Grand Total" + String(fa.Cols - 1, Chr(9)) + Format(tot, fmt)
    Else
        For k = 0 To colg - 1
            s = s + fa.TextMatrix(j, k) + Chr(9)
        Next
        s = s + "Total " + fa.TextMatrix(j, colg)
        For k = colg + 2 To fa.Cols - 1
            s = s + Chr(9)
        Next
        s = s + Chr(9) + Format(tot, fmt)
    End If
    
    ' add it to table
    fa.AddItem s, i
    fa.RowData(i) = -10000 + colg
    
    ' set color
    fa.Row = i
    fa.Col = colg
    fa.ColSel = fa.Cols - 1
    fa.CellBackColor = clr
    fa.CellFontBold = True
End Sub

Sub TotalOn(colg%, colv%, clr&, fmt$)
    Dim i%, j%, k%, ins%
    Dim tot#
    
    ' initialize
    tot = 0
    i = fa.FixedRows
    j = -1
    
    ' loop through rows
    Do
        ' get out when done
TryNext:
        If i >= fa.Rows Then Exit Do
        
        ' skip over subtotals
        If fa.RowData(i) < 0 Then
            i = i + 1
            GoTo TryNext
        End If
        
        ' see if we need to insert a subtotal here
        ins = False
        If j > -1 Then
            For k = 0 To colg
                If fa.TextMatrix(i, k) <> fa.TextMatrix(j, k) Then
                    ins = True
                    Exit For
                End If
            Next
        End If
            
        ' yes, go ahead and do it
        If ins Then
            TotalInsert tot, colg, i, j, clr, fmt
            i = i + 1
            tot = 0
            j = -1
        
        ' no, keep adding total
        Else
            tot = tot + fa.ValueMatrix(i, colv)
            j = i
            i = i + 1
        End If
    Loop
    
    ' anything left over?
    If j >= 1 Then
        i = fa.Rows
        TotalInsert tot, colg, i, j, clr, fmt
    End If
End Sub


Private Sub Check1_Click()
    If Check1 Then
        fa.MergeCells = 2
    Else
        fa.MergeCells = 0
    End If
End Sub

Private Sub Check2_Click()
    If Check2 Then
        fa.OutlineBar = 1
    Else
        fa.OutlineBar = 0
        fa.ColWidth(0) = 0
    End If

End Sub

Private Sub Check3_Click()
    SortAndShow
End Sub


Private Sub DoFillData()
    
    Dim i%
    
    For i = fa.FixedRows To fa.Rows - 1
      fa.TextMatrix(i, 1) = Region()
      fa.TextMatrix(i, 2) = Product()
      fa.TextMatrix(i, 3) = Employee()
      fa.TextMatrix(i, 4) = Sales()
    Next
    
    Combo1.AddItem "flexSum"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTSum)
    Combo1.AddItem "flexPercent"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTPercent)
    Combo1.AddItem "flexMin"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTMin)
    Combo1.AddItem "flexMax"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTMax)
    Combo1.AddItem "flexAverage"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTAverage)
    Combo1.AddItem "flexVar"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTVar)
    Combo1.AddItem "flexStd"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTStd)
    Combo1.AddItem "flexCount"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTCount)
    Combo1.AddItem "flexNone"
    Combo1.ItemData(Combo1.NewIndex) = Val(flexSTNone)
    Combo1.ListIndex = 0

    ' change fonts on title Row
    fa.Cell(flexcpFontBold, 0, 0, 0, fa.Cols - 1) = True
    fa.Cell(flexcpForeColor, 0, 0, 0, fa.Cols - 1) = vbYellow
End Sub

Private Sub DoRemoveTotals()
    Dim i%
    
    i = 1
    While i < fa.Rows
        If fa.RowData(i) < 0 Then
            fa.RemoveItem i
        Else
            i = i + 1
        End If
    Wend
End Sub



Private Function Employee() As String
    Select Case (Rnd * 1000) Mod 3
        Case 0: Employee = "Mary"
        Case 1: Employee = "Sarah"
        Case Else: Employee = "Paula"
    End Select
End Function

Private Sub Combo1_Click()

   Select Case Combo1.ItemData(Combo1.ListIndex)
      Case flexSTPercent
         m_Fmt = "###.00"
      Case flexSTCount
         m_Fmt = "##0"
      Case Else
         m_Fmt = "$(#,###.00)"
   End Select
   
   SortAndShow
   
End Sub

Private Sub Command3_Click()
    MousePointer = 11
    fa.Redraw = False
    fa.SubtotalPosition = 1 - fa.SubtotalPosition
    SortAndShow
    fa.Redraw = True
    MousePointer = 0
End Sub


Private Sub fa_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SortAndShow
End Sub

Private Sub fa_AfterMoveColumn(ByVal Col As Long, Position As Long)
    SortAndShow
End Sub


Private Sub fa_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' don't edit subtotals or merged stuff
    If fa.MergeCells And Col < 3 Then
        Cancel = True
        Exit Sub
    End If
    If fa.IsSubtotal(Row) Then
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub fa_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    
    ' don't move values
    If Col = 4 Then Position = 4

End Sub

Private Sub Form_Load()
    
    ' hide outline column
    fa.ColWidth(0) = 0
        
    ' initialize columns
    Dim i%
    For i = 1 To 3
        fa.ColWidth(i) = 1800
        fa.MergeCol(i) = True
    Next i
    fa.ColComboList(1) = "North|South|East|West"
    fa.ColComboList(2) = "VS-OCX|VSFlex|VSView"
    fa.ColComboList(3) = "Mary|Paula|Sarah"
    fa.ColFormat(4) = "#,###.##"
    
    ' create some data
    DoFillData
    SortAndShow
    
    ' start with auto totals and outline
    Check2.Value = 1
    Check3.Value = 1

End Sub

Private Function Product() As String
    Select Case (Rnd * 1000) Mod 4
        Case 0: Product = "VS-OCX"
        Case 1: Product = "VSView"
        Case Else: Product = "VSFlex"
    End Select
End Function

Private Function Region() As String
    Select Case (Rnd * 1000) Mod 4
        Case 0: Region = "North"
        Case 1: Region = "East"
        Case 2: Region = "South"
        Case Else: Region = "West"
    End Select
End Function

Private Function Sales() As String
    Sales = 5000 + Rnd * 10000
End Function

Private Sub Form_Resize()
    Dim X!
    X = ScaleWidth - 2 * fa.left
    If X > 0 Then fa.Width = X
    X = ScaleHeight - fa.left - fa.top
    If X > 0 Then fa.Height = X
End Sub


