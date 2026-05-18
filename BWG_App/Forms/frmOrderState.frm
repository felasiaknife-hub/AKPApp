VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmCustomerOrderState 
   Caption         =   "Order Status"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   120
      TabIndex        =   11
      Top             =   1380
      Width           =   3315
      Begin VB.OptionButton Option1 
         Caption         =   "ยกเลิกการระงับออเดอร์"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2835
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ระงับออเดอร์"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   2835
      End
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   7035
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   3540
      TabIndex        =   5
      Top             =   1380
      Width           =   4215
      Begin VB.CommandButton cmdSave 
         Caption         =   "บันทึก"
         Height          =   555
         Left            =   2340
         TabIndex        =   8
         Top             =   240
         Width           =   1515
      End
      Begin WasteManagment.ctlDate ctlDate1 
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label lblStartDate 
         Caption         =   "วันที่เริ่มระงับออเดอร์"
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   -120
      TabIndex        =   2
      Top             =   6360
      Width           =   8055
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   6540
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsgOrderState 
      Height          =   3915
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   7755
      _cx             =   13679
      _cy             =   6906
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   -60
      TabIndex        =   0
      Top             =   -120
      Width           =   7995
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "บันทึกการระงับออเดอร์"
         Height          =   255
         Left            =   540
         TabIndex        =   4
         Top             =   300
         Width           =   4395
      End
   End
   Begin VB.Label Label3 
      Caption         =   "ลูกค้า"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1140
      Width           =   435
   End
End
Attribute VB_Name = "frmCustomerOrderState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim cIndex As Integer
Dim DefaultMsg As String
Dim CanAccess As Boolean, CanUpdate As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    Dim tSql As String, OrderStatus As Integer
    Dim RS As New ADODB.Recordset
    Dim OrderDesc As String, OrderDateStart As String, OrderDateStop As String
    OrderDesc = InputBox("ใส่หมายเหตุในการระงับหรือยกเลิกการระงับออเดอร์", "หมายเหตุ", DefaultMsg)
    If MsgBox("ยืนยันการบันทึกข้อมูล", vbYesNo) = vbNo Then Exit Sub
    If cIndex = 0 Then OrderStatus = 2   ' Break
    If cIndex = 1 Then OrderStatus = 1 ' Open
    DBConn.BeginTrans
    Select Case OrderStatus
        Case 1 'ตรวจสอบดูว่ามีการระงับออเดอร์อยู่หรือไม่ ถ้าไม่มีก็ให้ออกจาก sub นี้ไปเลย ถ้ามีให้อัพเดทวันที่สิ้นสุดการระงับออเดอร์และสถานะ active=0 และบันทึกการยกเลิกการระงับออเดอร์
            tSql = "SELECT * FROM tbCustomerOrderBreak " & vbCrLf
            tSql = tSql & "WHERE CustomerID = '" & curCustomer_.ID & "' AND OrderStatus = 2" & vbCrLf
            tSql = tSql & "AND OrderStatusActive = 1 AND OrderDateStart <= '" & ctlDate1.ValueYMD & "'"
            Set RS = New ADODB.Recordset
            RS.open tSql, DBConn, adOpenDynamic, adLockOptimistic
            If RS.EOF Then
                MsgBox "ลูกค้ารายนี้ไม่มีการระงับออเดอร์ หรือระบุวันที่น้อยกว่าวันที่ที่มีการระงับออเดอร์, ไม่สามารถยกเลิกการระงับออเดอร์ได้", vbInformation
                DBConn.RollbackTrans
                Exit Sub
            Else
                RS!OrderDateStop = ctlDate1.ValueYMD
                RS!OrderStatusActive = 0
                RS.Update
            End If
        Case 2 'ตรวจสอบดูว่า มีการยกเลิกการระงับออเดอร์อยู่หรือไม่ ถ้าไม่มีก็ให้เพิ่มการระงับออเดอร์ไปตามปกติ ถ้ามีให้อัพเดทวันที่สิ้นสุดการยกเลิกการระงับออเดอร์และสถานะ active=0 ด้วย และบันทึกการงะงับออเดอร์
            tSql = "SELECT * FROM tbCustomerOrderBreak " & vbCrLf
            tSql = tSql & "WHERE CustomerID = '" & curCustomer_.ID & "' AND OrderStatus = 1" & vbCrLf
            tSql = tSql & "AND OrderStatusActive = 1 AND (OrderDateStart <= '" & ctlDate1.ValueYMD & "' OR OrderDateStart > '" & ctlDate1.ValueYMD & "')"
            Set RS = New ADODB.Recordset
            RS.open tSql, DBConn, adOpenDynamic, adLockOptimistic
            If Not RS.EOF Then
                If ctlDate1.ValueYMD < RS!OrderDateStart Then
                    MsgBox "มีรายการยกเลิกการระงับออเดอร์ ที่มีวันที่มากกว่าที่ที่ระบุ เพราะเฉพาะนั้นต้องระบุวันที่ที่ต้องการระงับมากกกว่าวันนที่ที่ยกเลิกการระงับออเดอร์", vbInformation
                    DBConn.RollbackTrans
                    Exit Sub
                Else
                    RS!OrderDateStop = ctlDate1.ValueYMD
                    RS!OrderStatusActive = 0
                    RS.Update
                End If
            End If
            Set RS = Nothing
    End Select
    
    tSql = "SELECT * FROM tbCustomerOrderBreak " & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & curCustomer_.ID & "' AND OrderStatus = " & OrderStatus & vbCrLf
    tSql = tSql & "AND OrderStatusActive = 1 AND OrderDateStart >= '" & ctlDate1.ValueYMD & "'"
    Set RS = New ADODB.Recordset
    RS.open tSql, DBConn, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
        MsgBox "ข้อมูลซ้ำ หรือมีรายการ [" & DefaultMsg & "] อยู่ก่อนแล้ว", vbInformation
        DBConn.RollbackTrans
        Exit Sub
    Else
        RS.AddNew
        RS!CustomerID = curCustomer_.ID
        RS!OrderStatus = OrderStatus
        RS!OrderStatusActive = 1
        RS!OrderDateStart = ctlDate1.ValueYMD
        RS!OrderDateStop = ""
        RS!OrderBreakDesc = OrderDesc
        RS.Update
    End If
    DBConn.CommitTrans
    Set RS = Nothing
    InitialData
    Exit Sub
ErrHandler:
    DBConn.RollbackTrans
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    SetGrid
    Option1_Click 0
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        lblStartDate.Caption = "วันที่เริ่มระงับออเดอร์"
        DefaultMsg = "ระงับออเดอร์"
    ElseIf Index = 1 Then
        lblStartDate.Caption = "วันที่ยกเลิกการระงับออเดอร์"
        DefaultMsg = "ยกเลิกการระงับออเดอร์"
    End If
    cIndex = Index
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
    On Error GoTo ErrD
    Dim tmpRec As New ADODB.Recordset
    Dim x%, sqlStr As String
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    txtCustomer.Text = curCustomer_.CustomerName
    InitialData
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub InitialData()
    Dim tSql As String, OrderStatus As String
    Dim RS As New ADODB.Recordset
    tSql = "SELECT * FROM tbCustomerOrderBreak" & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & curCustomer_.ID & "'" & vbCrLf
    tSql = tSql & "ORDER BY OrderDateStart, OrderStatusActive"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    vsgOrderState.Clear flexcpData
    Do While Not RS.EOF
        vsgOrderState.Rows = RS.RecordCount + 2
        With vsgOrderState
            Select Case RS!OrderStatus
                Case 1
                    OrderStatus = "ยกเลิกการระงับออเดอร์"
                Case 2
                    OrderStatus = "ระงับออเดอร์"
            End Select
            .TextMatrix(RS.AbsolutePosition, 0) = RS.AbsolutePosition
            .TextMatrix(RS.AbsolutePosition, 1) = OrderStatus
            .TextMatrix(RS.AbsolutePosition, 2) = RS!OrderDateStart
            .TextMatrix(RS.AbsolutePosition, 3) = RS!OrderDateStop
            .TextMatrix(RS.AbsolutePosition, 4) = RS!OrderBreakDesc
        End With
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub SetGrid()
    With vsgOrderState
        .Cols = 5
        .Rows = 20
        .FixedCols = 0
        .ColWidth(0) = 700
        .ColWidth(1) = 1800
        .ColWidth(2) = 1300
        .ColWidth(3) = 1300
        .ColWidth(4) = 2000
        .TextMatrix(0, 0) = "ลำดับ"
        .TextMatrix(0, 1) = "สถานะ"
        .TextMatrix(0, 2) = "วันที่เริ่มสถานะ"
        .TextMatrix(0, 3) = "วันที่สิ้นสุดสถานะ"
        .TextMatrix(0, 4) = "หมายเหตุ"
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .FixedAlignment(4) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColDataType(0) = flexDTString
        .ColDataType(1) = flexDTString
        .ColDataType(2) = flexDTString
        .ColDataType(3) = flexDTString
        .ColDataType(4) = flexDTString
        .AllowUserResizing = flexResizeColumns
    End With
End Sub
