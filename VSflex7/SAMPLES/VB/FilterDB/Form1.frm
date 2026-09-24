VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsFlex7.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid Demo: FilterData"
   ClientHeight    =   5970
   ClientLeft      =   1710
   ClientTop       =   1380
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   5610
   Begin VB.Frame frmDAO 
      Caption         =   "DAO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   40
      TabIndex        =   4
      Top             =   2820
      Width           =   5500
      Begin VB.CommandButton btnRefreshDAO 
         Caption         =   "&Refresh Data"
         Height          =   315
         Left            =   3615
         TabIndex        =   5
         Top             =   240
         Width           =   1590
      End
      Begin VB.Data Data1 
         Caption         =   "DAO Data Source"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Width           =   2600
      End
   End
   Begin VB.Frame frmADO 
      Caption         =   "ADO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   40
      TabIndex        =   2
      Top             =   0
      Width           =   5500
      Begin VB.CommandButton btnRefreshADO 
         Caption         =   "&Refresh Data"
         Height          =   315
         Left            =   3615
         TabIndex        =   3
         Top             =   240
         Width           =   1590
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   960
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
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
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "ADO DataSource"
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
   End
   Begin VSFlex7DAOCtl.VSFlexGrid fgDAO 
      Bindings        =   "Form1.frx":0000
      Height          =   2415
      Left            =   60
      TabIndex        =   1
      Top             =   3480
      Width           =   5340
      _cx             =   9419
      _cy             =   4260
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
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   1
      VirtualData     =   -1  'True
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
   Begin VSFlex7Ctl.VSFlexGrid fgADO 
      Bindings        =   "Form1.frx":0014
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   5400
      _cx             =   9525
      _cy             =   3731
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
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   1
      VirtualData     =   -1  'True
      DataMember      =   ""
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

Private Sub btnRefreshADO_Click()
    Adodc1.Refresh
End Sub

Private Sub btnRefreshDAO_Click()
    Data1.Refresh
End Sub

Private Sub fgADO_AfterDataRefresh()
    With fgADO
        .ColEditMask(3) = "99/99/99"
        .ColEditMask(4) = "99/99/99"
    End With
End Sub

Private Sub fgADO_FilterData(ByVal Row As Long, ByVal Col As Long, Data As String, ByVal SavingToDB As Boolean, WantThisCol As Boolean)
    
    ' we are interested only in columns 3 and 4
    If Col >= 3 And Col <= 4 Then
    
        ' keep getting this column
        WantThisCol = True
        
        ' format data going out to the recordset
        If SavingToDB Then
            If Len(Data) = 8 Then
                Data = Left(Data, 2) & Mid(Data, 4, 2) & Right(Data, 2)
            End If
            
        ' format data coming in from the recordset
        Else
            If Len(Data) = 6 Then
                Data = Left(Data, 2) & "/" & Mid(Data, 3, 2) & "/" & Right(Data, 2)
            End If
        End If
    End If

End Sub

Private Sub fgDAO_AfterDataRefresh()
    With fgDAO
        .ColEditMask(3) = "99/99/99"
        .ColEditMask(4) = "99/99/99"
    End With
End Sub

Private Sub fgDAO_FilterData(ByVal Row As Long, ByVal Col As Long, Data As String, ByVal SavingToDB As Boolean, WantThisCol As Boolean)
    
    ' use same handler as ADO grid
    fgADO_FilterData Row, Col, Data, SavingToDB, WantThisCol

End Sub

Private Sub Form_Load()

    ' set data sources
    With Adodc1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\test.mdb"
        .RecordSource = "NameData"
        .Refresh
    End With
    With Me.Data1
        .DatabaseName = App.Path & "\test.mdb"
        .RecordSource = "NameData"
    End With
    
    ' initialize grids
    With fgADO
        .AllowUserResizing = flexResizeBoth
        .Editable = flexEDKbdMouse
    End With
    With fgDAO
        .AllowUserResizing = flexResizeBoth
        .Editable = flexEDKbdMouse
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With frmADO
        Dim h!
        h = ScaleHeight / 2 - .Height - 2 * .Left
        .Move .Left, .Top, ScaleWidth - 2 * .Left
        fgADO.Move .Left, .Top + .Height + .Left, .Width, h
        frmDAO.Move .Left, fgADO.Top + fgADO.Height + .Left, .Width, .Height
        fgDAO.Move .Left, frmDAO.Top + frmDAO.Height + .Left, .Width, h
    End With
End Sub
