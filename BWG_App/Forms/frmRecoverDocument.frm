VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRecoverDocument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "กู้คืนเอกสาร"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   3435
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   8595
      _cx             =   15161
      _cy             =   6059
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   435
      Left            =   4260
      TabIndex        =   2
      Top             =   240
      Width           =   1155
   End
   Begin VB.TextBox txtDocNo 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   300
      Width           =   2715
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton cmdRecover 
      Caption         =   "Recovery"
      Height          =   495
      Left            =   6540
      TabIndex        =   4
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "เลขที่เอกสาร"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmRecoverDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DocType As String
Dim TmpValue As String
Dim ActValue As String

Private Type UpdateQuery
    SqlString As String
End Type
Dim colUpdateQuery() As UpdateQuery
Dim StrTable As String, StrWHERE As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRecover_Click()
On Error GoTo ErrD
    Dim tSql As String, i As Integer
    Dim uQuery As String
    
    If MsgBox("ยืนยันการกู้ข้อมูล " & DocType & " เลขที่ " & txtDocNo.Text, vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    For i = 0 To UBound(colUpdateQuery) - 1
        If colUpdateQuery(i).SqlString <> "" Then
            If uQuery = "" Then
                uQuery = colUpdateQuery(i).SqlString
            Else
                uQuery = uQuery & ", " & colUpdateQuery(i).SqlString
            End If
        End If
    Next i
    tSql = "UPDATE " & StrTable & " SET " & uQuery & StrWHERE
    DBConn.Execute tSql
    MsgBox "Update Completed!!", vbInformation
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub cmdSearch_Click()
    Dim RS As New ADODB.Recordset
    Dim tSql As String, tSqlCond As String
    Dim i As Integer
    TmpValue = ""
    ActValue = ""
    Select Case DocType
        Case "Quotation"    'ใบเสนอราคา
            StrTable = "tbQuotationHeader "
            StrWHERE = " WHERE QuotationNo = '" & Trim(txtDocNo) & "'"
            tSqlCond = "SELECT * FROM " & StrTable & StrWHERE
        Case "JobNo"          'ออเดอร์
            StrTable = "tbJobDataHeader"
            StrWHERE = " WHERE JobNo = '" & Trim(txtDocNo) & "'"
            tSqlCond = "SELECT * FROM " & StrTable & StrWHERE
        Case "Menifest"         'ใบเมนิเฟส
            StrTable = "tbMenifestHeader"
            StrWHERE = " WHERE DocNo = '" & Trim(txtDocNo) & "'"
            tSqlCond = "SELECT * FROM " & StrTable & StrWHERE
        Case "BillingNo"      'ใบ Inv
            StrTable = "tbBillingHeader"
            StrWHERE = " WHERE BillingNo = '" & Trim(txtDocNo) & "'"
            tSqlCond = "SELECT * FROM " & StrTable & StrWHERE
        Case "InvNo"             'ใบเสร็จรับเงิน
            StrTable = "tbPrintInvHeader"
            StrWHERE = " WHERE InvNo = '" & Trim(txtDocNo) & "'"
            tSqlCond = "SELECT * FROM " & StrTable & StrWHERE
    End Select
    If Trim(tSqlCond) = "" Then Exit Sub
    tSql = tSqlCond
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If RS.RecordCount <= 0 Then MsgBox "ไม่พบข้อมูลที่ค้นหา", vbInformation: Exit Sub
    SetGrid
    Do While Not RS.EOF
            With VSFlexGrid1
                For i = 0 To RS.Fields.Count - 1
                    If InStr(LCase(RS.Fields(i).Name), "company") <> 0 Or InStr(LCase(RS.Fields(i).Name), "no") <> 0 Or InStr(LCase(RS.Fields(i).Name), "cancel") <> 0 Or InStr(LCase(RS.Fields(i).Name), "close") <> 0 Then
                        .AddItem RS.Fields(i).Name & vbTab & RS.Fields(i).Value
                    End If
                Next i
            End With
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub Form_Load()
    ReDim colUpdateQuery(1)
    SetGrid
End Sub

Private Sub SetGrid()
    With VSFlexGrid1
        .Clear
        .Cols = 2
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1
        .ColWidth(0) = 2000
        .ColWidth(1) = .Width - .ColWidth(0) - 350
        .TextMatrix(0, 0) = "Column List"
        .TextMatrix(0, 1) = "Value"
        .ColAlignment(0) = flexAlignLeftBottom
        .ColAlignment(1) = flexAlignLeftBottom
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub txtDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then Exit Sub
    ActValue = VSFlexGrid1.TextMatrix(Row, 0) & "='" & VSFlexGrid1.TextMatrix(Row, 1) & "'"
    If ActValue <> TmpValue Then StringSQL ActValue
    TmpValue = ""
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True: Exit Sub
    If VSFlexGrid1.TextMatrix(Row, 0) = "isCancel" Or VSFlexGrid1.TextMatrix(Row, 0) = "isCanceled" Or VSFlexGrid1.TextMatrix(Row, 0) = "isClose" Or VSFlexGrid1.TextMatrix(Row, 0) = "isClosed" Then
        Cancel = False
        If TmpValue = "" Or TmpValue <> VSFlexGrid1.TextMatrix(Row, 0) & "='" & VSFlexGrid1.TextMatrix(Row, 1) & "'" Then
            TmpValue = VSFlexGrid1.TextMatrix(Row, 0) & "='" & VSFlexGrid1.TextMatrix(Row, 1) & "'"
        End If
    Else
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub StringSQL(ByVal UpdateString As String)
    Dim i As Integer
    For i = 0 To UBound(colUpdateQuery) - 1
        If colUpdateQuery(i).SqlString = UpdateString Then
            UpdateString = ""
            Exit Sub
        End If
    Next i
    
    ReDim Preserve colUpdateQuery(UBound(colUpdateQuery) + 1)
    If UpdateString <> "" Then colUpdateQuery(UBound(colUpdateQuery) - 2).SqlString = UpdateString
        
End Sub
