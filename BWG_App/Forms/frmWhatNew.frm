VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmWhatNew 
   Caption         =   "What New?"
   ClientHeight    =   6900
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   10365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   6375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      _cx             =   18230
      _cy             =   11245
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
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
      ShowComboButton =   1
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
End
Attribute VB_Name = "frmWhatNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With VSFlexGrid1
        .Clear
        .Cols = 2
        .Rows = 20
        .FixedCols = 0
        .TextMatrix(0, 0) = "Date"
        .TextMatrix(0, 1) = "Detail"
        .ColWidth(0) = 2000
        .ColWidth(1) = .Width - .ColWidth(0)
    End With
    
    Dim RS As New ADODB.Recordset
    Dim tSql  As String
    tSql = "SELECT * FROM tbWasteManagementRevision" & vbCrLf
    tSql = tSql & "WHERE Program = '" & App.Title & "'"
    tSql = tSql & "ORDER BY IndexUpdate, DateUpdate"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If RS.RecordCount > 0 Then
        VSFlexGrid1.Rows = RS.RecordCount + 10
        Do Until RS.EOF
            With VSFlexGrid1
                .TextMatrix(RS.AbsolutePosition, 0) = Trim("" & RS!DateUpdate)
                .TextMatrix(RS.AbsolutePosition, 1) = Trim("" & RS!DetailUpdate)
            End With
            RS.MoveNext
        Loop
    Else
        VSFlexGrid1.Rows = 30
    End If
    VSFlexGrid1.ScrollBars = flexScrollBarBoth
    Set RS = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With VSFlexGrid1
        .Height = Me.Height - cmdExit.Height - 750
        .Width = Me.Width - 225
        .ColWidth(0) = 2000
        .ColWidth(1) = .Width - (.ColWidth(0) * 1.16)
    End With
    cmdExit.Top = Me.Height - cmdExit.Height - 650
    cmdExit.Left = Me.Width - cmdExit.Width - 350
'    Me.Caption = "Me.Height = " & Me.Height & " Me.Width = " & Me.Width & " cmdExit.Top = " & cmdExit.Top & " cmdexit.Left  = " & cmdExit.Left
End Sub
