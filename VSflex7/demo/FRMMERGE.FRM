VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmMerge 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Summary Sheet"
   ClientHeight    =   6000
   ClientLeft      =   1575
   ClientTop       =   1635
   ClientWidth     =   8760
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   8760
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VSFlexGrid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   7
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1020
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
      Begin VB.CheckBox chkPictures 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Pictures"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         TabIndex        =   6
         Top             =   420
         Width           =   1245
      End
      Begin VB.CheckBox chkMergeCells 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Merge Cells"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         TabIndex        =   5
         Top             =   675
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Drag && Drop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   4
         Top             =   1020
         Width           =   1095
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plain Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1125
      TabIndex        =   2
      Top             =   1770
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      DragIcon        =   "frmMerge.frx":0000
      Height          =   4635
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   4845
      _cx             =   5080
      _cy             =   5080
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   0
      ForeColorFixed  =   65280
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   12632256
      GridColor       =   32896
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   420
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "^Region            |^Product       |^Type       |>Sales            "
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin VB.Line Line1 
      Index           =   1
      X1              =   660
      X2              =   5820
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   540
      Shape           =   3  'Circle
      Top             =   660
      Width           =   135
   End
   Begin VB.Image imgConvert 
      Height          =   300
      Index           =   8
      Left            =   900
      Picture         =   "frmMerge.frx":030A
      Top             =   300
      Width           =   1710
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "See summary as:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Index           =   0
      Left            =   450
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Image flag 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   750
      Picture         =   "frmMerge.frx":13EF
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image flag 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   135
      Picture         =   "frmMerge.frx":16F9
      Top             =   4095
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image flag 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   1950
      Picture         =   "frmMerge.frx":1A03
      Top             =   4095
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image flag 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   1275
      Top             =   4125
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUMREC = 10

Dim Mydata(NUMREC) As MYDATA_TYPE

Private Type MYDATA_TYPE
    Country As String
    Product As String
    Type As String
    Sales As String
End Type

Private Sub DoGetData()

    Dim i%
    
    i = i + 1
    Mydata(i).Country = "Germany"
    Mydata(i).Product = "Gold"
    Mydata(i).Type = "Export"
    Mydata(i).Sales = "1,267"

    i = i + 1
    Mydata(i).Country = "Germany"
    Mydata(i).Product = "Gold"
    Mydata(i).Type = "Import"
    Mydata(i).Sales = "167"


    i = i + 1
    Mydata(i).Country = "Germany"
    Mydata(i).Product = "Silver"
    Mydata(i).Type = "Import"
    Mydata(i).Sales = "167"


    i = i + 1
    Mydata(i).Country = "Canada"
    Mydata(i).Product = "Silver"
    Mydata(i).Type = "Import"
    Mydata(i).Sales = "70"

    i = i + 1
    Mydata(i).Country = "Canada"
    Mydata(i).Product = "Silver"
    Mydata(i).Type = "Export"
    Mydata(i).Sales = "70"

    i = i + 1
    Mydata(i).Country = "Canada"
    Mydata(i).Product = "Gold"
    Mydata(i).Type = "Export"
    Mydata(i).Sales = "70"

    i = i + 1
    Mydata(i).Country = "USA"
    Mydata(i).Product = "Silver"
    Mydata(i).Type = "Import"
    Mydata(i).Sales = "70"

    i = i + 1
    Mydata(i).Country = "USA"
    Mydata(i).Product = "Silver"
    Mydata(i).Type = "Export"
    Mydata(i).Sales = "70"

    i = i + 1
    Mydata(i).Country = "USA"
    Mydata(i).Product = "Gold"
    Mydata(i).Type = "Export"
    Mydata(i).Sales = "70"


    i = i + 1
    Mydata(i).Country = "USA"
    Mydata(i).Product = "Gold"
    Mydata(i).Type = "Import"
    Mydata(i).Sales = "70"

End Sub


Private Sub chkMergeCells_Click()
    If fa.MergeCells = 0 Then
        fa.MergeCells = 2
    Else
        fa.MergeCells = 0
    End If
End Sub

Private Sub chkPictures_Click()
   Dim i%

   fa.Row = 0
   
   For i = 0 To 3
      ' find "Region" column
      fa.Col = i
      If Trim(fa) = "Region" Then Exit For
   Next i
   
   If chkPictures Then
      DoAddFlagsToColumn fa.Col
   Else
      fa.Cell(flexcpPicture, 1, fa.Col, fa.Rows - 1) = flag(0)
   End If

End Sub

Private Sub Command1_Click()
   MsgBox "Drag and drop column headers and see the data automatically rearranged"
End Sub

Private Sub DoAddFlagsToColumn(Col)
   Dim i%
   
   fa.Col = Col
   
   ' align pictures
   fa.Cell(flexcpPictureAlignment, 1, Col) = flexPicAlignLeftCenter
   
   For i = fa.FixedRows To fa.Rows - 1
      fa.CellPictureAlignment = 1
      Select Case fa.Cell(flexcpText, i, Col)
         Case "USA": fa.Cell(flexcpPicture, i, Col) = flag(1)
         Case "Canada": fa.Cell(flexcpPicture, i, Col) = flag(2)
         Case "Germany": fa.Cell(flexcpPicture, i, Col) = flag(3)
      End Select
   Next i
   
   

End Sub

Private Sub DoFillData()
   
   Dim i%
   
   fa.Rows = 1
   
   ' from structure to FlexArray
   For i = 1 To UBound(Mydata)
      
      fa.AddItem Mydata(i).Country & Chr$(9) & Mydata(i).Product & Chr$(9) & Mydata(i).Type & Chr$(9) & Format$(Mydata(i).Sales, "##.00")
      
      ' negatives in red
      If Val(Mydata(i).Sales) < 0 Then fa.Cell(flexcpForeColor, i, 3) = RGB(255, 0, 0)
    
    Next

End Sub
Private Sub DoInitialSettings()

    Dim i%

    fa.ColAlignment(0) = flexAlignLeftCenter

    ' change fonts on title Row
    fa.Cell(flexcpFontSize, 0, 0, 0, fa.Cols - 1) = 12
    fa.Cell(flexcpAlignment, 0, 0, 0, fa.Cols - 1) = flexAlignCenterCenter
    
    ' merge only first 3 columns
    For i = 0 To 3
        fa.MergeCol(i) = True
    Next i

End Sub

Private Sub DoSort()
    
    fa.Col = 0
    fa.ColSel = fa.Cols - 1
    fa.Sort = flexSortGenericAscending
    
End Sub

Private Sub fa_AfterMoveColumn(ByVal Col As Long, Position As Long)
    DoSort
End Sub

Private Sub fa_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = 3 Then Position = 3
End Sub

Private Sub Form_Load()
    DoInitialSettings
    DoGetData
    DoFillData
    DoSort
End Sub

Private Sub Option1_Click(Index As Integer)
   
   ' enable and disable controls
   chkPictures.Enabled = Not chkPictures.Enabled
   chkMergeCells.Enabled = Not chkMergeCells.Enabled
   Command1.Enabled = Not Command1.Enabled
   Frame1.Enabled = Not Frame1.Enabled

   fa.Redraw = False
   chkPictures.Value = Index
   chkMergeCells.Value = Index
   fa.Redraw = True

End Sub

