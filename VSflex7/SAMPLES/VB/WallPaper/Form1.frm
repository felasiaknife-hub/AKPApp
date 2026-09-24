VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "FlexGrid WallPaper"
   ClientHeight    =   4725
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   9675
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   4560
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=NorthWind"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "NorthWind"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Customers"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   60
      List            =   "Form1.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4140
      Width           =   1755
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   60
      Pattern         =   "*.jpg;*.gif;*.wmf"
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Align           =   4  'Align Right
      Bindings        =   "Form1.frx":0128
      Height          =   4725
      Left            =   1935
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _cx             =   13652
      _cy             =   8334
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
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
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FormatCells     =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label1 
      Caption         =   "Align"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "WallPaper"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    fg.WallPaperAlignment = Combo1.ListIndex
End Sub

Private Sub File1_Click()
    On Error Resume Next
    fg.WallPaper = LoadPicture(File1.Path & "\" & File1.List(File1.ListIndex))
End Sub

Private Sub Form_Load()
    File1.Path = App.Path
    With fg
        .BackColor = &HF4F4F4
        .MergeCells = flexMergeSpill
        .AllowUserFreezing = flexFreezeBoth
        .GridLines = flexGridNone
    End With
End Sub

Private Sub Form_Resize()
    With File1
        fg.Width = ScaleWidth - .Width - 2 * .Left
    End With
End Sub
