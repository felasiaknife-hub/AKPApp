VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTruckMonitor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   7350
   ClientLeft      =   2715
   ClientTop       =   3375
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTick 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7860
      TabIndex        =   2
      Text            =   "60"
      Top             =   6960
      Width           =   375
   End
   Begin VB.Timer TimerGlobal 
      Interval        =   1000
      Left            =   3000
      Top             =   60
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   60
   End
   Begin VB.CommandButton cmdView 
      Height          =   315
      Left            =   1560
      Picture         =   "frmTruckMonitor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   315
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   6315
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   9795
      _cx             =   17277
      _cy             =   11139
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
      BackColorAlternate=   16777152
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   79495169
      CurrentDate     =   40809
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   9720
      Picture         =   "frmTruckMonitor.frx":0171
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   3
      Top             =   60
      Width           =   390
   End
   Begin VB.Label lblReflesh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Auto left refresh s."
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8655
      TabIndex        =   6
      Top             =   7020
      Width           =   1275
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "System..."
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   7020
      Width           =   7515
   End
End
Attribute VB_Name = "frmTruckMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ServerTime As Date
Dim TickValue As Integer
Dim WorkDate As String

Private Sub cmdView_Click()
    WorkDate = DTPicker1.Year + 543 & "/" & Format(DTPicker1.Month, "00") & "/" & Format(DTPicker1.Day, "00")
    VSFlexGrid1.Clear 2, 1
    TruckMonitor
    TickValue = 0
End Sub

Private Sub DTPicker1_Change()
    DTPicker1.CustomFormat = "dd/MM/" & Trim(Year(DTPicker1.Value))
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    DTPicker1.UpDown = False
    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = "dd/MM/" & Trim(Year(DTPicker1.Value))
    DoEvents
    Call SetGrid
    RegionFromImage Picture1, Picture1, GetPixel(Picture1.hdc, 0, 0)
    MakeFormRounded Me, 25, , 2
    TimerGlobal.Enabled = True
    Timer1.Enabled = False
    Timer1.Interval = 1000
    ServerTime = GetDateServer
    lblReflesh.Caption = "Auto left reflesh " & txtTick.Text & " s."
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   FormDrag Me
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub

Private Sub TruckMonitor()
On Error GoTo ErrD
    Dim RS As New ADODB.Recordset
    Dim tSql As String
   '*** NEW SQL
    tSql = "SELECT     TOP (100) PERCENT '" & WorkDate & "' AS WorkDate, SubTypeID, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(dbo.tbTruckTotalAvailable.TruckTotal, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              From dbo.vw_JobData_Truck_Count" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID = dbo.tbTruckTotalAvailable.SubTypeID) AND ('" & WorkDate & "' >= dbo.tbTruckTotalAvailable.TruckAvailableDate) AND (WorkDate = '" & WorkDate & "') AND (ISNULL(isCancel, 'N') <> 'Y')" & vbCrLf
    tSql = tSql & "                              ORDER BY WorkDate DESC), 0) AS TruckTotal, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckNotAvailable, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              From dbo.tbTruckManage" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID = dbo.tbTruckTotalAvailable.SubTypeID) AND ('" & WorkDate & "' BETWEEN TruckNotAvailableDateStart AND TruckNotAvailableDateEnd)), 0)" & vbCrLf
    tSql = tSql & "                      AS TruckNotAvailable, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS vw_JobData_Truck_Count_1" & vbCrLf
    tSql = tSql & "                              WHERE     (CompanyID = '776EDA5830F744A6B3551348D851BC22') AND (SubTypeID = dbo.tbTruckTotalAvailable.SubTypeID) AND (WorkDate = '" & WorkDate & "') AND" & vbCrLf
    tSql = tSql & "                                                    (ISNULL(isCancel, 'N') <> 'Y')), 0) AS AKPTruckUsed, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS vw_JobData_Truck_Count_2" & vbCrLf
    tSql = tSql & "                              WHERE     (CompanyID = 'C451AF07C9A44B4282D6E08B14AC0195') AND (SubTypeID = dbo.tbTruckTotalAvailable.SubTypeID) AND (WorkDate = '" & WorkDate & "') AND" & vbCrLf
    tSql = tSql & "                                                    (ISNULL(isCancel, 'N') <> 'Y')), 0) AS BWCTruckUsed, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS vw_JobData_Truck_Count_3" & vbCrLf
    tSql = tSql & "                              WHERE     (CompanyID NOT IN ('776EDA5830F744A6B3551348D851BC22', 'C451AF07C9A44B4282D6E08B14AC0195')) AND" & vbCrLf
    tSql = tSql & "                                                    (SubTypeID = dbo.tbTruckTotalAvailable.SubTypeID) AND (WorkDate = '" & WorkDate & "') AND (ISNULL(isCancel, 'N') <> 'Y')), 0) AS BWGTruckUsed," & vbCrLf
    tSql = tSql & "                      SubTypeDesc" & vbCrLf
    tSql = tSql & "From dbo.tbTruckTotalAvailable" & vbCrLf
    tSql = tSql & "GROUP BY SubTypeDesc, SubTypeID, TruckAvailableDate" & vbCrLf
    tSql = tSql & "ORDER BY SubTypeDesc"

    '*** OLD SQL
'    tSql = "SELECT     TOP (100) PERCENT CONVERT(DATETIME, CAST(YEAR(WorkDate) - 543 AS VARCHAR(4)) + '/' + CAST(MONTH(WorkDate) AS VARCHAR(2)) + '/' + CAST(DAY(WorkDate)" & vbCrLf
'    tSql = tSql & "                      AS VARCHAR(2)), 111) AS Expr1, SubTypeID, ISNULL" & vbCrLf
'    tSql = tSql & "                          ((SELECT     TOP (1) ISNULL(TruckTotal, 0) AS Expr1" & vbCrLf
'    tSql = tSql & "                              FROM         dbo.tbTruckTotalAvailable AS tbTruckTotalAvailable_2" & vbCrLf
'    tSql = tSql & "                              Where (SubTypeID = dbo.vw_JobData_Truck_Count.SubTypeID) And (dbo.vw_JobData_Truck_Count.WorkDate >= TruckAvailableDate)" & vbCrLf
'    tSql = tSql & "                              ORDER BY TruckAvailableDate DESC), 0) AS TruckTotal, ISNULL" & vbCrLf
'    tSql = tSql & "                          ((SELECT     TOP(1) SUM(ISNULL(TruckNotAvailable, 0))" & vbCrLf
'    tSql = tSql & "                              FROM         dbo.tbTruckManage AS tbTruckManage_1" & vbCrLf
'    tSql = tSql & "                              WHERE     (SubTypeID = dbo.vw_JobData_Truck_Count.SubTypeID) AND (dbo.vw_JobData_Truck_Count.WorkDate BETWEEN TruckNotAvailableDateStart AND" & vbCrLf
'    tSql = tSql & "                                                    TruckNotAvailableDateEnd)), 0) AS TruckNotAvailable, ISNULL" & vbCrLf
'    tSql = tSql & "                          ((SELECT     SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
'    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS V1" & vbCrLf
'    tSql = tSql & "                              WHERE     (CompanyID = '776EDA5830F744A6B3551348D851BC22') AND (SubTypeID = dbo.vw_JobData_Truck_Count.SubTypeID) AND" & vbCrLf
'    tSql = tSql & "                                                    (WorkDate = dbo.vw_JobData_Truck_Count.WorkDate) AND (ISNULL(isCancel, 'N') <> 'Y')), 0) AS AKPTruckUsed, ISNULL" & vbCrLf
'    tSql = tSql & "                          ((SELECT     SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
'    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS V1" & vbCrLf
'    tSql = tSql & "                              WHERE     (CompanyID = 'C451AF07C9A44B4282D6E08B14AC0195') AND (SubTypeID = dbo.vw_JobData_Truck_Count.SubTypeID) AND" & vbCrLf
'    tSql = tSql & "                                                    (WorkDate = dbo.vw_JobData_Truck_Count.WorkDate) AND (ISNULL(isCancel, 'N') <> 'Y')), 0) AS BWCTruckUsed, ISNULL" & vbCrLf
'    tSql = tSql & "                          ((SELECT     SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
'    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS V1" & vbCrLf
'    tSql = tSql & "                              WHERE     (CompanyID NOT IN ( '776EDA5830F744A6B3551348D851BC22', 'C451AF07C9A44B4282D6E08B14AC0195')) AND (SubTypeID = dbo.vw_JobData_Truck_Count.SubTypeID) AND" & vbCrLf
'    tSql = tSql & "                                                    (WorkDate = dbo.vw_JobData_Truck_Count.WorkDate) AND (ISNULL(isCancel, 'N') <> 'Y')), 0) AS BWGTruckUsed, TruckTypeDesc, SubTypeDesc" & vbCrLf
'    tSql = tSql & "From dbo.vw_JobData_Truck_Count" & vbCrLf
'    tSql = tSql & "WHERE     (WorkDate = '" & WorkDate & "') AND (ISNULL(isCancel, 'N') <> 'Y')" & vbCrLf
'    tSql = tSql & "GROUP BY WorkDate, TruckTypeDesc, SubTypeDesc, TruckTypeID, SubTypeID" & vbCrLf
'    tSql = tSql & "ORDER BY TruckTypeDesc, SubTypeDesc"
    
    Screen.MousePointer = 11
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If RS.RecordCount = 0 Then Exit Sub
    Timer1.Enabled = True
    VSFlexGrid1.Rows = RS.RecordCount + 2
    Do While Not RS.EOF
        With VSFlexGrid1
            .TextMatrix(RS.AbsolutePosition + 1, 0) = RS.AbsolutePosition
            .TextMatrix(RS.AbsolutePosition + 1, 1) = RS!SubTypeDesc
            .TextMatrix(RS.AbsolutePosition + 1, 2) = RS!TruckTotal - RS!TruckNotAvailable
            .TextMatrix(RS.AbsolutePosition + 1, 3) = RS!BWGTruckUsed
            .TextMatrix(RS.AbsolutePosition + 1, 4) = RS!BWCTruckUsed
            .TextMatrix(RS.AbsolutePosition + 1, 5) = RS!AKPTruckUsed
            .TextMatrix(RS.AbsolutePosition + 1, 6) = (RS!TruckTotal - RS!TruckNotAvailable) - (RS!BWGTruckUsed + RS!BWCTruckUsed + RS!AKPTruckUsed)
            If (RS!TruckTotal - RS!TruckNotAvailable) - (RS!BWGTruckUsed + RS!BWCTruckUsed + RS!AKPTruckUsed) < 0 Then
                .Cell(flexcpForeColor, RS.AbsolutePosition + 1, .Cols - 1, RS.AbsolutePosition + 1, .Cols - 1) = vbRed
                .Cell(flexcpFontBold, RS.AbsolutePosition + 1, .Cols - 1, RS.AbsolutePosition + 1, .Cols - 1) = True
            ElseIf (RS!TruckTotal - RS!TruckNotAvailable) - (RS!BWGTruckUsed + RS!BWCTruckUsed + RS!AKPTruckUsed) = 0 Then
                .Cell(flexcpForeColor, RS.AbsolutePosition + 1, .Cols - 1, RS.AbsolutePosition + 1, .Cols - 1) = vbBlue
                .Cell(flexcpFontBold, RS.AbsolutePosition + 1, .Cols - 1, RS.AbsolutePosition + 1, .Cols - 1) = True
            Else
                .Cell(flexcpForeColor, RS.AbsolutePosition + 1, .Cols - 1, RS.AbsolutePosition + 1, .Cols - 1) = &H8000&
                .Cell(flexcpFontBold, RS.AbsolutePosition + 1, .Cols - 1, RS.AbsolutePosition + 1, .Cols - 1) = True
            End If
            DoEvents
        End With
        RS.MoveNext
    Loop
    VSFlexGrid1.SubtotalPosition = flexSTBelow
'    VSFlexGrid1.Subtotal flexSTClear
    VSFlexGrid1.Subtotal flexSTSum, -1, 2, "#,", vbYellow, vbBlue, True
    VSFlexGrid1.Subtotal flexSTSum, -1, 3, "#,", vbYellow, vbBlue, True
    VSFlexGrid1.Subtotal flexSTSum, -1, 4, "#,", vbYellow, vbBlue, True
    VSFlexGrid1.Subtotal flexSTSum, -1, 5, "#,", vbYellow, vbBlue, True
    VSFlexGrid1.Subtotal flexSTSum, -1, 6, "#,", vbYellow, vbBlue, True
    
    RS.Close
    Set RS = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub SetGrid()
Dim FixCol As Long
    With VSFlexGrid1
        .Cols = 7
        .Rows = 10
        .FixedRows = 2
        .FixedCols = 1
        .ColWidth(0) = 500
        .ColWidth(1) = 3000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1500
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "Truck Type"
        .TextMatrix(0, 2) = "Truck Balance"
        .TextMatrix(0, 3) = "Truck Used"
        .TextMatrix(0, 4) = "Truck Used"
        .TextMatrix(0, 5) = "Truck Used"
        .TextMatrix(0, 6) = "Truck Available"
        .TextMatrix(1, 0) = "No"
        .TextMatrix(1, 1) = "Truck Type"
        .TextMatrix(1, 2) = "Truck Balance"
        .TextMatrix(1, 3) = "BWG"
        .TextMatrix(1, 4) = "BWC"
        .TextMatrix(1, 5) = "AKP"
        .TextMatrix(1, 6) = "Truck Available"
        .MergeCells = flexMergeFixedOnly
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(6) = True
        .MergeRow(0) = True
        .MergeRow(1) = True
        For FixCol = 0 To .Cols - 1
            .FixedAlignment(FixCol) = flexAlignCenterCenter
        Next
        .Editable = flexEDNone
    End With
End Sub

Private Sub Timer1_Timer()
    TickValue = TickValue + 1
    If TickValue > Val(txtTick) Then TickValue = 0
    lblReflesh.Caption = "Auto left reflesh " & TickValue & " s."
    DoEvents
    
    ' *** Refresh Monitor
    If TickValue = Val(txtTick) Then
        cmdView_Click
    End If
    
End Sub

Private Sub TimerGlobal_Timer()
    ServerTime = DateAdd("s", 1, ServerTime)
    lblStatus.Caption = "System Date " & Format(ServerTime, "dd/MM/yyyy HH:mm:ss")
    DoEvents
End Sub

Private Sub txtTick_Validate(Cancel As Boolean)
    If IsNumeric(txtTick) = False Then
        MsgBox "Please input numeric only.", vbInformation
        txtTick.Text = "10"
        Exit Sub
    End If
    If IsNumeric(txtTick) Then
        If Val(txtTick) < 5 Then txtTick.Text = "10"
    End If
    lblReflesh.Caption = "Auto left reflesh " & txtTick.Text & " s."
End Sub
