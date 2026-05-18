VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTruckManage 
   Caption         =   "Truck Management"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdView 
      Height          =   315
      Left            =   1440
      Picture         =   "frmTruckManage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   315
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   8235
      _cx             =   14526
      _cy             =   8387
      _ConvInfo       =   1
      Appearance      =   2
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
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
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
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid2 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   5100
      Width           =   8235
      _cx             =   14526
      _cy             =   4471
      _ConvInfo       =   1
      Appearance      =   2
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
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   50921473
      CurrentDate     =   40787
   End
   Begin MSComCtl2.DTPicker DTDate 
      Height          =   315
      Left            =   6900
      TabIndex        =   4
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   50921473
      CurrentDate     =   40787
   End
End
Attribute VB_Name = "frmTruckManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isEnDate As Boolean
Dim WorkDate As String

Private Sub cmdView_Click()
    Dim msg As String
    WorkDate = DTPicker1.Year + 543 & "/" & Format(DTPicker1.Month, "00") & "/" & Format(DTPicker1.Day, "00")
    VSFlexGrid1.Clear 1, 1
    msg = LoadMaster(WorkDate)
    If msg <> "" Then
        MsgBox msg
        msg = ""
    End If
End Sub

Private Sub DTDate_Change()
    DTDate.CustomFormat = "dd/MM/" & Trim(Year(DTDate.Value))
    VSFlexGrid2.Text = Format(DTDate.Value, "dd/MM/") & Year(DTDate.Value)
End Sub

Private Sub DTPicker1_Change()
    DTPicker1.CustomFormat = "dd/MM/" & Trim(Year(DTPicker1.Value))
End Sub

Private Sub Form_Load()
    Dim Localization As New CLSCulture
    DTPicker1.Value = Now
    DTPicker1.UpDown = False
    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = "dd/MM/" & Trim(Year(DTPicker1.Value))
    
    DTDate.Value = Now
    DTDate.UpDown = False
    DTDate.Format = dtpCustom
    DTDate.CustomFormat = "dd/MM/" & Trim(Year(DTPicker1.Value))
    DoEvents
    
    DTDate.CausesValidation = False
    
    If Val(Format(DTDate.Value, "yyyy")) = DTDate.Year Then
        isEnDate = True
    Else
        isEnDate = False
    End If
    
    Me.Icon = frmMain.Icon
    Call SetGrid
    Call LoadMaster
    Call GetTruckHeader
    Call GetTruckNotAvailable
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    VSFlexGrid1.Width = Me.Width - 225
'    VSFlexGrid1.Height = Me.Height - VSFlexGrid2.Height
    VSFlexGrid2.Width = Me.Width - 225
    VSFlexGrid2.Height = Me.Height - VSFlexGrid1.Height - 100
End Sub

Private Sub SetGrid()
Dim Col As Integer
    With VSFlexGrid1
        .Cols = 7
        .Rows = 18
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 5000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 2000
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "TruckTypeID"
        .TextMatrix(0, 2) = "ประเภทรถ"
        .TextMatrix(0, 3) = "จำนวนรถทั้งหมด"
        .TextMatrix(0, 4) = "จำนวนรถที่ใช้ได้"
        .TextMatrix(0, 5) = "จำนวนรถที่ใช้ไม่ได้"
        .TextMatrix(0, 6) = "วันที่เริ่มใช้งานได้"
        For Col = 0 To .Cols - 1
            .FixedAlignment(Col) = flexAlignCenterCenter
            If Col > 2 Then .ColAlignment(Col) = flexAlignCenterCenter
        Next
        .Editable = flexEDKbdMouse
    End With
    With VSFlexGrid2
        .Rows = 8
        .FixedCols = 1
        .FixedRows = 2
        .ColWidth(0) = 800
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .TextMatrix(1, 0) = "No"
        .TextMatrix(1, 1) = "StartDate"
        .TextMatrix(1, 2) = "EndDate"
        .ColFormat(1) = "dd/MM/" & Year(Now)
        .ColFormat(2) = "dd/MM/" & Year(Now)
        For Col = 0 To .Cols - 1
            .FixedAlignment(Col) = flexAlignCenterCenter
            .ColAlignment(Col) = flexAlignCenterCenter
        Next
        .Editable = flexEDKbdMouse
    End With
    DTDate.Visible = False
End Sub

Private Sub GetTruckHeader()
Dim RS As ADODB.Recordset
Dim tSql As String, Col As Integer

    tSql = "SELECT     TOP (100) PERCENT dbo.tbTruckType.TruckTypeID, dbo.tbTruckType.TruckTypeDesc, dbo.tbTruckSubType.SubTypeID," & vbCrLf
    tSql = tSql & "                      dbo.tbTruckSubType.SubTypeDesc , dbo.tbTruckType.isActive" & vbCrLf
    tSql = tSql & "FROM         dbo.tbTruckType INNER JOIN" & vbCrLf
    tSql = tSql & "                      dbo.tbTruckSubType ON dbo.tbTruckType.TruckTypeID = dbo.tbTruckSubType.TruckTypeID" & vbCrLf
    tSql = tSql & "WHERE     (dbo.tbTruckType.isActive = 'Y')  AND (dbo.tbTruckSubType.isActive = 'Y') AND (dbo.tbTruckType.TruckTypeDesc NOT IN ('รถโรงงาน'))" & vbCrLf
    tSql = tSql & "ORDER BY dbo.tbTruckType.TruckTypeDesc, dbo.tbTruckSubType.SubTypeDesc"

    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If RS.RecordCount = 0 Then Exit Sub
    VSFlexGrid2.Cols = RS.RecordCount + 3
    Col = 3
    Do While Not RS.EOF
        VSFlexGrid2.TextMatrix(0, 0) = ""
        VSFlexGrid2.TextMatrix(0, 1) = ""
        VSFlexGrid2.TextMatrix(0, Col) = RS!SubTypeID
        VSFlexGrid2.TextMatrix(1, Col) = RS!SubTypeDesc
        If Col >= 3 Then VSFlexGrid2.ColWidth(Col) = 1200
        RS.MoveNext
        Col = Col + 1
    Loop
    RS.Close
    Set RS = Nothing
    VSFlexGrid2.RowHeight(0) = 0
End Sub

Private Sub GetTruckNotAvailable()
On Error Resume Next
Dim RS As ADODB.Recordset
Dim tSql As String, Row As Integer, Col As Integer
Dim TruckID_Header As String
    
    tSql = "SELECT     TOP (100) PERCENT dbo.tbTruckType.TruckTypeID, dbo.tbTruckType.TruckTypeDesc, dbo.tbTruckSubType.SubTypeID," & vbCrLf
    tSql = tSql & "                      dbo.tbTruckSubType.SubTypeDesc , dbo.tbTruckType.isActive" & vbCrLf
    tSql = tSql & "FROM         dbo.tbTruckType INNER JOIN" & vbCrLf
    tSql = tSql & "                      dbo.tbTruckSubType ON dbo.tbTruckType.TruckTypeID = dbo.tbTruckSubType.TruckTypeID" & vbCrLf
    tSql = tSql & "WHERE     (dbo.tbTruckType.isActive = 'Y') AND (dbo.tbTruckSubType.isActive = 'Y') AND (dbo.tbTruckType.TruckTypeDesc NOT IN ('รถโรงงาน'))" & vbCrLf
    tSql = tSql & "ORDER BY dbo.tbTruckType.TruckTypeDesc, dbo.tbTruckSubType.SubTypeDesc"
    
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    TruckID_Header = ""
    If RS.RecordCount = 0 Then Exit Sub
    Do While Not RS.EOF
        If TruckID_Header = "" Then
            TruckID_Header = "MAX(CASE WHEN SubTypeID = '" & Trim(RS!SubTypeID) & "' THEN ISNULL(TruckNotAvailable, 0) ELSE 0 END) AS '[" & RS!SubTypeDesc & "]'" & vbCrLf
        Else
            TruckID_Header = TruckID_Header & ", MAX(CASE WHEN SubTypeID = '" & Trim(RS!SubTypeID) & "' THEN ISNULL(TruckNotAvailable, 0) ELSE 0 END) AS '[" & RS!SubTypeDesc & "]'" & vbCrLf
        End If
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    
    If Trim(TruckID_Header) = "" Then Exit Sub
    tSql = "SELECT Row_Number() OVER(ORDER BY TruckNotAvailableDateStart) AS RowNo, TruckNotAvailableDateStart, TruckNotAvailableDateEnd," & vbCrLf
    tSql = tSql & TruckID_Header
    tSql = tSql & "From tbTruckManage" & vbCrLf
    tSql = tSql & "WHERE TruckNotAvailableDateStart IS NOT NULL" & vbCrLf
    tSql = tSql & "GROUP BY TruckNotAvailableDateStart, TruckNotAvailableDateEnd"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    TruckID_Header = ""
    If RS.RecordCount = 0 Then Exit Sub
    VSFlexGrid2.Rows = RS.RecordCount + 3
    Row = 2
    Do While Not RS.EOF
        With VSFlexGrid2
            '.TextMatrix(Row, 0) = RS.AbsolutePosition
            For Col = 0 To RS.Fields.Count - 1
                If IsDate(RS.Fields(Col).Value) Then
                    .TextMatrix(Row, Col) = Day(RS.Fields(Col).Value) & "/" & Month(RS.Fields(Col).Value) & "/" & Year(RS.Fields(Col).Value) - 543
                Else
                    .TextMatrix(Row, Col) = RS.Fields(Col).Value
                End If
            Next Col
        End With
        RS.MoveNext
        Row = Row + 1
    Loop
    RS.Close
    Set RS = Nothing
    
    With VSFlexGrid2
        .FixedCols = 1
        .FixedRows = 2
        .ColWidth(0) = 800
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .TextMatrix(1, 0) = "No"
        .TextMatrix(1, 1) = "StartDate"
        .TextMatrix(1, 2) = "EndDate"
        .ColFormat(1) = "dd/MM/" & Year(Now)
        .ColFormat(2) = "dd/MM/" & Year(Now)
        For Col = 0 To .Cols - 1
            .FixedAlignment(Col) = flexAlignCenterCenter
            .ColAlignment(Col) = flexAlignCenterCenter
        Next
        .Editable = flexEDKbdMouse
    End With
    DTDate.Visible = False
End Sub

Private Function LoadMaster(Optional ByVal AvailableDate As String = "") As String
    On Error GoTo ErrD
    Dim RS As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    Dim tSql As String, NotAvailableDate As String
    If AvailableDate <> "" Then
        AvailableDate = "'" & AvailableDate & "'"
        NotAvailableDate = ">= '" & AvailableDate & "'"
    End If
    If AvailableDate = "" Then
        AvailableDate = "''"
        NotAvailableDate = "''"
    End If
    LoadMaster = ""
    
    tSql = "SELECT DISTINCT SubTypeID, TruckTypeDesc, SubTypeDesc, isActive," & vbCrLf
    tSql = tSql & "                          (SELECT     MAX(TruckAvailableDate)" & vbCrLf
    tSql = tSql & "                            From dbo.tbTruckTotalAvailable" & vbCrLf
    tSql = tSql & "                            WHERE      (SubTypeID = dbo.vw_TruckManage.SubTypeID) AND (" & AvailableDate & " >= TruckAvailableDate)) AS TruckAvailableDate," & vbCrLf
    tSql = tSql & "                          ISNULL((SELECT TOP(1)    TruckTotal" & vbCrLf
    tSql = tSql & "                            FROM          dbo.tbTruckTotalAvailable AS tbTruckTotalAvailable_2" & vbCrLf
    tSql = tSql & "                            WHERE      (SubTypeID = dbo.vw_TruckManage.SubTypeID) AND (" & AvailableDate & " >= TruckAvailableDate) ORDER BY TruckAvailableDate DESC), 0) AS TruckTotal," & vbCrLf
    
    tSql = tSql & "                          ISNULL((SELECT TOP(1)    TruckTotal" & vbCrLf
    tSql = tSql & "                            FROM          dbo.tbTruckTotalAvailable AS tbTruckTotalAvailable_2" & vbCrLf
    tSql = tSql & "                            WHERE      (SubTypeID = dbo.vw_TruckManage.SubTypeID) AND (" & AvailableDate & " >= TruckAvailableDate) ORDER BY TruckAvailableDate DESC), 0) - " & vbCrLf
    tSql = tSql & "                          ISNULL((SELECT TOP(1)    TruckNotAvailable" & vbCrLf
    tSql = tSql & "                            FROM tbTruckManage WHERE (SubTypeID = dbo.vw_TruckManage.SubTypeID) AND (" & AvailableDate & " BETWEEN TruckNotAvailableDateStart AND TruckNotAvailableDateEnd)), 0) AS TruckAvailable, " & vbCrLf
    
    tSql = tSql & "                          ISNULL((SELECT TruckNotAvailable" & vbCrLf
    tSql = tSql & "                            FROM tbTruckManage WHERE (SubTypeID = dbo.vw_TruckManage.SubTypeID) AND (" & AvailableDate & " BETWEEN TruckNotAvailableDateStart AND TruckNotAvailableDateEnd)), 0) AS TruckNotAvailable" & vbCrLf
    tSql = tSql & "From dbo.vw_TruckManage" & vbCrLf
    tSql = tSql & "WHERE     (isActive = 'Y') AND (TruckTypeDesc NOT IN ('รถโรงงาน'))" & vbCrLf
    tSql = tSql & "ORDER BY TruckAvailableDate DESC, TruckTypeDesc, SubTypeDesc"
    
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    
    tSql = "SELECT * FROM tbTruckTotalAvailable "
    If AvailableDate <> "''" Then tSql = tSql & "WHERE TruckAvailableDate = " & AvailableDate & ""
    Set rsCheck = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rsCheck.RecordCount = 0 Then
        LoadMaster = "ไม่มีข้อมูลรถ คุณต้องทำการเพิ่มจำนวนรถในระบบก่อน โดยการระบุวันที่เริ่มต้นกำหนดจำนวนรถในระบบ และบันทึกจำนวนรถที่มีทั้งหมดของแต่ละประเภทรถ"
    End If
    VSFlexGrid1.Rows = RS.RecordCount + 1
    Do While Not RS.EOF
        With VSFlexGrid1
            .TextMatrix(RS.AbsolutePosition, 0) = Trim("" & RS.AbsolutePosition)
            .TextMatrix(RS.AbsolutePosition, 1) = Trim("" & RS!SubTypeID)
            .TextMatrix(RS.AbsolutePosition, 2) = Trim("" & RS!SubTypeDesc)
            .TextMatrix(RS.AbsolutePosition, 3) = Trim("" & RS!TruckTotal)
            .TextMatrix(RS.AbsolutePosition, 4) = Trim("" & RS!TruckAvailable)
            .TextMatrix(RS.AbsolutePosition, 5) = Trim("" & RS!TruckNotAvailable)
            If Replace(AvailableDate, "'", "") <> "" Then
                If Replace(AvailableDate, "'", "") <> Trim("" & RS!TruckAvailableDate) Then
                    .TextMatrix(RS.AbsolutePosition, 6) = Replace(AvailableDate, "'", "")
                Else
                    .TextMatrix(RS.AbsolutePosition, 6) = Trim("" & RS!TruckAvailableDate)
                End If
            Else
                .TextMatrix(RS.AbsolutePosition, 6) = Trim("" & RS!TruckAvailableDate)
            End If
        End With
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    Exit Function
ErrD:
    MsgBox err.Description
    Exit Function
End Function

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrD
    Dim RS As New ADODB.Recordset
    Dim tSql  As String, iCol As Integer
    WorkDate = DTPicker1.Year + 543 & "/" & Format(DTPicker1.Month, "00") & "/" & Format(DTPicker1.Day, "00")
    With VSFlexGrid1
        Select Case Col
            Case 3
                If IsNumeric(.TextMatrix(Row, Col)) = False Then .TextMatrix(Row, Col) = "0"
                If Trim(.TextMatrix(Row, 1)) = "" Then Exit Sub
                If Trim(.TextMatrix(Row, .Cols - 1)) <> "" Then WorkDate = .TextMatrix(Row, .Cols - 1)
                tSql = "SELECT * FROM tbTruckTotalAvailable " & vbCrLf
                tSql = tSql & "WHERE TruckAvailableDate = '" & WorkDate & "'" & vbCrLf
                tSql = tSql & "AND SubTypeID = '" & .TextMatrix(Row, 1) & "'"
                RS.open tSql, DBConn, adOpenKeyset, adLockOptimistic
                If RS.RecordCount = 0 Then
                    RS.AddNew
                End If
                
                RS!SubTypeID = .TextMatrix(Row, 1)
                RS!SubTypeDesc = .TextMatrix(Row, 2)
                RS!TruckTotal = .TextMatrix(Row, 3)
                RS!TruckAvailableDate = WorkDate
                RS.Update
                RS.Close
                Set RS = Nothing
                
                .TextMatrix(Row, 4) = .TextMatrix(Row, 3) - IIf(IsNumeric(.TextMatrix(Row, 5)), .TextMatrix(Row, 5), 0)
        End Select
    End With
    
    If VSFlexGrid1.Rows = Row Then VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case 0, 1, 2, 4, 5
            Cancel = True
            Exit Sub
        Case 3
            If IsNumeric(VSFlexGrid1.TextMatrix(Row, Col)) = False Then
                MsgBox "Input Numeric only"
                Cancel = True
            End If
    End Select
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    With VSFlexGrid1
        If KeyCode = vbKeyReturn Then
            If .RowSel = .Rows - 1 Then
                If .ColSel = .Cols - 1 Then
                    .Row = 1
                    .Col = 1
                Else
                    .Col = .ColSel + 1
                    .Row = 1
                End If
            Else
                .Row = .RowSel + 1
            End If
        End If
    End With
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrD
    Dim RS As New ADODB.Recordset
    Dim tSql  As String, iCol As Integer
    With VSFlexGrid2
        Select Case Col
            Case Is >= 3
                If IsNumeric(.TextMatrix(Row, Col)) = False Then .TextMatrix(Row, Col) = "0"
                If Trim(.TextMatrix(Row, 1)) = "" Then Exit Sub
                If Trim(.TextMatrix(Row, 2)) = "" Then Exit Sub
                For iCol = 3 To .Cols - 1
                    tSql = "SELECT * FROM tbTruckManage WHERE TruckNotAvailableDateStart = '" & Format(DateAdd("yyyy", 543, .TextMatrix(Row, 1)), "yyyy/MM/dd") & "' AND TruckNotAvailableDateEnd = '" & Format(DateAdd("yyyy", 543, .TextMatrix(Row, 2)), "yyyy/MM/dd") & "' " & vbCrLf
                    tSql = tSql & "AND SubTypeID = '" & .TextMatrix(0, iCol) & "'"
                    RS.open tSql, DBConn, adOpenKeyset, adLockOptimistic
                    If RS.RecordCount = 0 Then
                        RS.AddNew
                    End If
                    
                    RS!SubTypeID = .TextMatrix(0, iCol)
                    RS!TruckTotal = 0
                    RS!TruckAvailable = 0
                    RS!TruckNotAvailable = IIf(IsNumeric(.TextMatrix(Row, iCol)), .TextMatrix(Row, iCol), 0)
                    RS!TruckNotAvailableDateStart = Format(DateAdd("yyyy", 543, .TextMatrix(Row, 1)), "yyyy/MM/dd")
                    RS!TruckNotAvailableDateEnd = Format(DateAdd("yyyy", 543, .TextMatrix(Row, 2)), "yyyy/MM/dd")
                    RS.Update
                    RS.Close
                    Set RS = Nothing
                Next iCol
                                
                For iCol = 3 To .Cols - 1
                    tSql = "UPDATE tbTruckManage SET " & vbCrLf
                    tSql = tSql & "TruckTotal = (SELECT TruckTotal FROM tbTruckTotalAvailable T2 WHERE T2.SubTypeID = tbTruckManage.SubTypeID AND  tbTruckManage.TruckNotAvailableDateStart >= T2.TruckAvailableDate)" & vbCrLf
                    tSql = tSql & ", TruckAvailable = (SELECT TruckTotal FROM tbTruckTotalAvailable T2 WHERE T2.SubTypeID = tbTruckManage.SubTypeID AND  tbTruckManage.TruckNotAvailableDateStart >= T2.TruckAvailableDate) - TruckNotAvailable" & vbCrLf
                    tSql = tSql & "WHERE SubTypeID = '" & .TextMatrix(0, iCol) & "'" & vbCrLf
                    tSql = tSql & "AND TruckNotAvailableDateStart = '" & Format(DateAdd("yyyy", 543, .TextMatrix(Row, 1)), "yyyy/MM/dd") & "'" & vbCrLf
                    tSql = tSql & "AND TruckNotAvailableDateEnd = '" & Format(DateAdd("yyyy", 543, .TextMatrix(Row, 2)), "yyyy/MM/dd") & "'"
                    DBConn.Execute tSql
                Next iCol
                cmdView_Click
        End Select
    End With
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub VSFlexGrid2_EnterCell()
    Dim SelCurDate As String
    With VSFlexGrid2
        Select Case .Col
            Case 1, 2
                If isEnDate = True Then
                    If IsDate(.Text) = True Then
                        SelCurDate = .Text
                    Else
                        SelCurDate = Format(Now, "dd/MM/" & Year(Now))
                    End If
                Else
                    If IsDate(.Text) = True Then
                        SelCurDate = Format(.Text, "dd/MM/" & Year(Now) + 543)
                    Else
                        SelCurDate = Format(Now, "dd/MM/" & Year(Now) + 543)
                    End If
                End If
                DTDate.Visible = True
                DTDate.Value = SelCurDate
                DTDate.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight + 20
                DTDate.ZOrder
            Case Else
                DTDate.Visible = False
        End Select
        Select Case .Row
            Case Is >= 2
                
        End Select
    End With
End Sub

Private Sub VSFlexGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
    If VSFlexGrid2.Row = VSFlexGrid2.Rows - 1 Then VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
End Sub

