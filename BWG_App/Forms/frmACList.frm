VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmACList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account List"
   ClientHeight    =   4680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   4035
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   9855
      _cx             =   17383
      _cy             =   7117
      _ConvInfo       =   -1
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
      BackColorAlternate=   12640511
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
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   7380
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmACList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private callerForm_ As Form
Private AccountNo As String

Public Sub SetCallerForm(selFrm As Form)
    Set callerForm_ = selFrm
End Sub

Private Sub CancelButton_Click()
        callerForm_.SetAccountNo ""
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    SetGrid
    LoadAccount
End Sub

Private Sub SetGrid()
    With VSFlexGrid1
        .Clear
        .Cols = 5
        .Rows = 20
        .ColWidth(0) = 500
        .ColWidth(1) = 1000     'BankCode
        .ColWidth(2) = 1500     'BranchName
        .ColWidth(3) = 1500     'AccountNo
        .ColWidth(4) = 5000     'AccountName
        
        .ColAlignment(0) = flexAlignCenterBottom
        .ColAlignment(1) = flexAlignCenterBottom
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignCenterBottom
        .ColAlignment(4) = flexAlignLeftBottom
        
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Bank"
        .TextMatrix(0, 2) = "Branch"
        .TextMatrix(0, 3) = "Account No."
        .TextMatrix(0, 4) = "Account Name"
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
End Sub

Public Sub LoadAccount(Optional ByVal Index As String = "")
    Dim tSql  As String
    Dim RS As New ADODB.Recordset
    tSql = "SELECT * FROM Com_vwBankAccountSearch "
    If Index <> "" Then tSql = tSql & "WHERE RowIndex = '" & Index & "'"
    tSql = tSql & "ORDER BY RowIndex, BankID, BranchID, AccountNo"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If RS.RecordCount <= 0 Then Exit Sub
    VSFlexGrid1.Rows = 1
    Do While Not RS.EOF
        VSFlexGrid1.AddItem Trim("" & RS!RowIndex) & vbTab & Trim("" & RS!BankCode) & vbTab & Trim("" & RS!BranchName) & vbTab & Trim("" & RS!AccountNo) & vbTab & Trim("" & RS!AccountName)
        RS.MoveNext
    Loop
    VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
End Sub

Private Sub OKButton_Click()
    SelectedData
End Sub

Private Sub VSFlexGrid1_DblClick()
    Dim Row&, Col&
    If VSFlexGrid1.Rows <= 0 Then Exit Sub
    If VSFlexGrid1.MouseRow <= 0 Then Exit Sub
    Row = VSFlexGrid1.MouseRow
    Col = VSFlexGrid1.MouseCol
    AccountNo = VSFlexGrid1.TextMatrix(Row, 3)
    callerForm_.SetAccountNo AccountNo
    Unload Me
End Sub

Public Sub SelectedData()
    Dim Row&, Col&
    With VSFlexGrid1
        If .RowSel <= 0 Then Exit Sub
        Row = .RowSel
        Col = 3
        AccountNo = .TextMatrix(Row, 3)
        callerForm_.SetAccountNo AccountNo
    End With
    Unload Me
End Sub
