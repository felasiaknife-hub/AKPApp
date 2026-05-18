VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmMenifestStatusListClose 
   Caption         =   "Manifest List"
   ClientHeight    =   9810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   11445
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid vsMenifestList 
      Height          =   6315
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   7695
      _cx             =   13573
      _cy             =   11139
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
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9810
      Left            =   0
      ScaleHeight     =   9780
      ScaleWidth      =   2625
      TabIndex        =   9
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frmMenifestStatusListClose.frx":0000
         Left            =   180
         List            =   "frmMenifestStatusListClose.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   2325
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   6
         Top             =   4140
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Top             =   4620
         Width           =   2295
      End
      Begin VB.TextBox txtKey2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Text            =   "*"
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   3720
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   1950
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   2610
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   18
         Top             =   120
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   2370
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   5550
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   5850
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   6150
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   3
         Left            =   180
         Top             =   6450
         Width           =   195
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   450
         TabIndex        =   15
         Top             =   5580
         Width           =   570
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พิมพ์ Manifest แล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   450
         TabIndex        =   13
         Top             =   6180
         Width           =   780
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   6480
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor, คนขับ, ทะเบียนรถ:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   2460
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบคุม :"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   3480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMenifestStatusListClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CopyObject As String
Private CopyColumns As Integer
Dim ClosedAll As Integer
Const slComboChargType = "ปริมาณผู้รับกำจัด|ปริมาณผู้ก่อกำเนิด|ไม่ได้กำหนด"
Dim CurMenifestID As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim RS As New ADODB.Recordset, RSRemark As New ADODB.Recordset
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim subStr As String
    sqlStr = "SELECT CustomerID, CustomerName, JobDetID, MenifestID, DocNo, WorkDate, TimeTableNo, WasteDataID, WasteName, " & vbCrLf
    sqlStr = sqlStr & "CASE WHEN ISNULL(isClosed,'N') <> 'Y' THEN CASE WHEN (SELECT ChargeWType FROM tbCustomer_TR_ChargeWType T2 WHERE T2.CustomerID = T1.CustomerID) = 'A' THEN 'ปริมาณผู้รับกำจัด' ELSE" & vbCrLf
    sqlStr = sqlStr & "                                           CASE WHEN (SELECT ChargeWType FROM tbCustomer_TR_ChargeWType T2 WHERE T2.CustomerID = T1.CustomerID) = 'B' THEN 'ปริมาณผู้ก่อกำเนิด' ELSE 'ไม่ได้กำหนด' END END ELSE" & vbCrLf
    sqlStr = sqlStr & "    CASE WHEN ISNULL(ChargeWType,'A') = 'A' THEN 'ปริมาณผู้รับกำจัด' ELSE" & vbCrLf
    sqlStr = sqlStr & "        CASE WHEN ISNULL(ChargeWType,'A') = 'B' THEN 'ปริมาณผู้ก่อกำเนิด' ELSE 'ปริมาณผู้รับกำจัด' END END END AS ChargeWType," & vbCrLf
    sqlStr = sqlStr & "isCanceled, isClosed, isPrinted, Status, ISNULL(DisposerWeight,0) AS DisposerWeight, ISNULL(CustWeight,0) AS CustWeight, CloseDate, CloseTime, DateArriveSite, TimeArriveSite" & vbCrLf
    sqlStr = sqlStr & "From Com_vwManifestWastAdd_Union T1" & vbCrLf

    frmMain.Stb1.Panels(1).Text = App.Title
    frmMain.Stb1.Panels(2).Text = ""
    Screen.MousePointer = 11
    If Trim(txtJobNo.Text) <> "" Then
        strCond = "WHERE TimeTableNo Like '%" & Trim(txtJobNo.Text) & "%'"
    Else
        strCond = "WHERE (WorkDate>='" & dtFrom.ValueYMD & "' AND WorkDate<='" & dtTo.ValueYMD & "') "
        If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
            strCond = strCond & " AND (DocNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' OR DocNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' OR CustomerName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
        End If
'        If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
'            strCond = strCond & " And (TransportName1 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or TransportName2 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%')"
'        End If
        Select Case cboJobStatus.ListIndex
            Case 1 'งานใหม่
                subStr = " ((ISNULL(isPrinted,'N')<>'Y') AND (ISNULL(isCanceled,'N')<>'Y') AND (ISNULL(isClosed,'N')<>'Y')) "
            Case 2 'พิมพ์ Menifest แล้ว
                subStr = " (ISNULL(isPrinted,'N')='Y' AND (ISNULL(isCanceled,'N')<>'Y') AND (ISNULL(isClosed,'N')<>'Y')) "
            Case 3 'ปิดงาน
                subStr = " (ISNULL(isClosed,'N')='Y' AND (ISNULL(isCanceled, 'N') <>'Y')) "
            Case 4 'ยกเลิก
                subStr = " (ISNULL(isCanceled,'N')='Y') "
        End Select
        If cboJobStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " AND " & subStr
            Else
                strCond = strCond & " WHERE " & subStr
            End If
        End If
    End If

    sqlStr = sqlStr & strCond
    sqlStr = SelectWithCompanyID(sqlStr, False)
    
     sqlStr = sqlStr & vbCrLf & "ORDER BY WorkDate, TimeTableNo, DocNo, WasteName"
    Set RS = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    If RS Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
        
    SetDefaults vsMenifestList
    With vsMenifestList
        .Rows = 1
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarComplete
        .SubtotalPosition = flexSTAbove
        Dim i%, j%, curMenifest$, AddedHeader As Boolean
        Dim RemarkList As String
        AddedHeader = False
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & RS.RecordCount & " รายการ"
        Do While Not RS.EOF
            DoEvents
            frmMain.Stb1.Panels(2).Text = "กำลังโหลดข้อมูลลงตาราง... กรุณารอซักครู่"
            If curMenifest = "" Then curMenifest = RS!docNo
            If curMenifest <> RS!docNo Then
                AddedHeader = False
                curMenifest = RS!docNo
            End If
            ' decide MenifestNo whether to add a subtotal
            If AddedHeader = False Then
                ' add an item, make it a subtotal
                sqlStr = "SELECT RemarkText FROM tbCustomer_TR_Remark WHERE CustomerID = '" & RS!CustomerID & "' ORDER BY RemarkNo"
                Set RSRemark = GetRS(sqlStr, adOpenForwardOnly, adLockOptimistic)
                If RSRemark.RecordCount > 0 Then
                    RemarkList = ""
                    Do While Not RSRemark.EOF
                        RemarkList = RemarkList & Format(RSRemark.AbsolutePosition, "0") & "." & RSRemark!RemarkText & ","
                        RSRemark.MoveNext
                    Loop
                Else
                    RemarkList = ""
                End If
                If Len(Trim(RemarkList)) > 0 Then RemarkList = Trim(Left(RemarkList, Len(RemarkList) - 1))
                .AddItem "" & vbTab & "H" & vbTab & RS!menifestID & vbTab & RS!JobDetID & vbTab & RS!docNo & vbTab & RS!CustomerName & vbTab & RS!ChargeWType & vbTab & "" & vbTab & "" & vbTab & Format(RS!DateArriveSite, "dd/MM/yyyy") & vbTab & RS!TimeArriveSite & vbTab & RemarkList & vbTab & RS!Status
                .IsSubtotal(.Rows - 1) = True
'                .RowOutlineLevel(.Rows - 1) = 0
                .Row = .Rows - 1
                .Col = 0
                .CellChecked = flexUnchecked
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                SetRowStyle RS!Status, .Rows - 1
                AddedHeader = True
                .AddItem "" & vbTab & "D" & vbTab & RS!menifestID & vbTab & RS!JobDetID & vbTab & RS!WasteName & vbTab & "" & vbTab & "" & vbTab & Format(RS!CustWeight, "0.000") & vbTab & Format(RS!DisposerWeight, "0.000")
                .Cell(flexcpText, .Rows - 1, 0) = ""
                .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = False
                SetRowStyle RS!Status, .Rows - 1
            Else
                'Add waste data at first row
                If curMenifest = RS!docNo Then
                    .AddItem "" & vbTab & "D" & vbTab & RS!menifestID & vbTab & RS!JobDetID & vbTab & RS!WasteName & vbTab & "" & vbTab & "" & vbTab & Format(RS!CustWeight, "0.000") & vbTab & Format(RS!DisposerWeight, "0.000")
                    .Cell(flexcpText, .Rows - 1, 0) = ""
                    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = False
                    SetRowStyle RS!Status, .Rows - 1
                End If
            End If
            If .Rows > 1 Then .Row = 0: .Col = 0
            RS.MoveNext
        Loop
        Set RS = Nothing
    End With
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(2).Text = ""
    ' do an autosize
'    vsMenifestList.AutoSize 0, 1, , 300
    
End Sub

Private Sub SetRowStyle(ByVal Status As String, ByVal Row As Long)
Dim curColor, i#
    Select Case Status
        Case "Canceled"
            curColor = Shape1(3).FillColor
        Case "Closed"
            curColor = Shape1(2).FillColor
        Case "Printed"
            curColor = Shape1(1).FillColor
        Case "New Job"
            curColor = Shape1(0).FillColor
        Case Else
            curColor = vbBlack
    End Select
    For i = 0 To vsMenifestList.Cols - 1
        vsMenifestList.Row = Row
        vsMenifestList.Col = i
        vsMenifestList.CellForeColor = curColor
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.ActiveControl.Container.Name = txtKey.Container.Name Then
            If Not TypeOf Me.ActiveControl Is CommandButton Then
                cmdSearch_Click
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrD
    With cboJobStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "งานใหม่"
        .AddItem "พิมพ์ Manifest แล้ว"
        .AddItem "ปิดงานแล้ว"
        .AddItem "ยกเลิก"
        .ListIndex = 0
    End With
    ' Hook Form
    Call WheelHook(Me.hWnd)
    
    SetDefaults vsMenifestList
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With vsMenifestList
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 200
        End If
        .Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Stb1.Panels(1).Text = App.Title
    frmMain.Stb1.Panels(2).Text = ""
    VisibleImage
'     UnHook Form
    Call WheelUnHook(Me.hWnd)
End Sub

Private Sub vsMenifestList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
Dim i&, hRow&, fRow&, lRow&
Dim TotalPW As Double, TotalCW As Double
    With vsMenifestList
        For i = Row To 1 Step -1
            If .TextMatrix(i, 1) = "H" Then
                hRow = i
                fRow = i + 1
                Exit For
            End If
        Next i
        For i = fRow To .Rows - 1
            If i = .Rows - 1 And .TextMatrix(i, 1) = "D" Then
                lRow = .Rows - 1
                Exit For
            End If
            If .TextMatrix(i, 1) = "H" Then
                lRow = i - 1
                Exit For
            End If
        Next i
        Select Case Col
            Case 7
                TotalCW = 0
                For i = fRow To lRow
                    TotalCW = TotalCW + IIf(IsNumeric(.TextMatrix(i, Col)), CDbl(.TextMatrix(i, Col)), 0)
                    .TextMatrix(hRow, Col) = Format(TotalCW, "0.000")
                Next i
                .TextMatrix(Row, Col) = Format(Replace(.TextMatrix(Row, Col), "..", "."), "0.000")
            
            Case 8
                TotalPW = 0
                For i = fRow To lRow
                    TotalPW = TotalPW + IIf(IsNumeric(Replace(.TextMatrix(i, Col), "..", ".")), CDbl(Replace(.TextMatrix(i, Col), "..", ".")), 0)
                    .TextMatrix(hRow, Col) = Format(TotalPW, "0.000")
                Next i
                .TextMatrix(Row, Col) = Format(Replace(.TextMatrix(Row, Col), "..", "."), "0.000")
            
            Case 9
                If Trim(.TextMatrix(Row, Col)) = "__/__/____" Then .TextMatrix(Row, Col) = "": Exit Sub
                If Trim(.TextMatrix(Row, Col)) = "" Then Exit Sub
                If IsDate(.TextMatrix(Row, Col)) = False Then
                    MsgBox "รูปแบบวันที่ไม่ถูกต้อง", vbCritical
                    .TextMatrix(Row, Col) = ""
                    Exit Sub
                Else
                    .TextMatrix(Row, Col) = DateValue(.TextMatrix(Row, Col))
                End If
            
            Case 10
                 If Trim(.TextMatrix(Row, Col)) = "__:__" Then .TextMatrix(Row, Col) = "": Exit Sub
                If Trim(.TextMatrix(Row, Col)) = "" Then Exit Sub
                If IsDate(.TextMatrix(Row, Col)) = False Then
                    MsgBox "รูปแบบเวลาไม่ถูกต้อง", vbCritical
                    .TextMatrix(Row, Col) = ""
                    Exit Sub
                Else
                    .TextMatrix(Row, Col) = Replace(.TextMatrix(Row, Col), ".", ":")
                End If
        End Select
    End With
End Sub

Private Sub vsMenifestList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case 1, 2, 3, 4, 5
            Cancel = True
            Exit Sub
        Case 0, 6, 9, 10, 11, 12
            If UCase(vsMenifestList.TextMatrix(Row, vsMenifestList.Cols - 1)) <> "PRINTED" Then Cancel = True
            If vsMenifestList.TextMatrix(Row, 1) = "D" Then Cancel = True
        Case 7, 8
            If vsMenifestList.TextMatrix(Row, 1) = "H" Then Cancel = True
    End Select
End Sub

Private Sub vsMenifestList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim Pt As POINTAPI
    
    ' get popup window position
    Pt.x = vsMenifestList.ColPos(Col) \ Screen.TwipsPerPixelX
    Pt.Y = (vsMenifestList.RowPos(Row) + vsMenifestList.RowHeight(Row)) \ Screen.TwipsPerPixelY
    ClientToScreen vsMenifestList.hWnd, Pt

    ' show message condition
    If Col = vsMenifestList.Cols - 1 Then
        If MsgBox("ยืนยันการปิดใบเมนิเฟส", vbInformation + vbYesNo) = vbYes Then
            '*** สั่งให้ปิดใบเมนิเฟส
            If Trim(vsMenifestList.TextMatrix(Row, 1)) = "" Then
                MsgBox "ไม่พบเอกสาร กรุณาลองใหม่"
                Exit Sub
            Else
                ClosedAll = 0
                CloseMF Row, 1
            End If
        End If
        Exit Sub
    End If
End Sub

Private Sub vsMenifestList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsMenifestList
        Select Case Col
            Case 9
                If Trim(.TextMatrix(Row, Col)) = "__/__/____" Then .TextMatrix(Row, Col) = "": Exit Sub
                If Trim(.TextMatrix(Row, Col)) = "" Then Exit Sub
            Case 10
                 If Trim(.TextMatrix(Row, Col)) = "__:__" Then .TextMatrix(Row, Col) = "": Exit Sub
                If Trim(.TextMatrix(Row, Col)) = "" Then Exit Sub
        End Select
    End With
End Sub

Private Sub vsMenifestList_DblClick()
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim r&, c&
    r = vsMenifestList.MouseRow
    c = vsMenifestList.MouseCol
    If r <= 0 Then Exit Sub
    If Trim(vsMenifestList.TextMatrix(r, 2)) = "" Then Exit Sub
    Set tmpCol = tmpS.MenifestSearch("MenifestID='" & vsMenifestList.TextMatrix(r, 2) & "'")
    If tmpCol.Count > 0 Then
        Load frmMenifest
        frmMenifest.SetDataMenifest tmpCol(1)
        frmMenifest.SSTab1.TabVisible(2) = False
        frmMenifest.Show 1
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub

Private Sub vsMenifestList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With vsMenifestList
            If .ColSel = .Cols - 1 Then
                If .RowSel = .Rows - 1 Then .RowSel = 0
                .Row = .RowSel + 1
                .Col = .FixedCols + 1
                .RowSel = .Row
                .ColSel = .Col
            Else
                .Col = .Col + 1
                .ColSel = .Col
            End If
        End With
    End If
End Sub

Private Sub vsMenifestList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If vsMenifestList.Rows <= 1 Then Exit Sub
    Dim Pont As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "Copy"
        AppendMenu hMenu, MF_STRING, 2, "Past"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP1"
        AppendMenu hMenu, MF_STRING, 3, "เลือกทั้งหมด"
        AppendMenu hMenu, MF_STRING, 4, "ไม่เลือกทั้งหมด"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP2"
        AppendMenu hMenu, MF_STRING, 5, "ปิดงานใบเมนิเฟสที่เลือก"
        AppendMenu hMenu, MF_STRING, 6, "ปิดงานใบเมนิเฟสที่เลือกทั้งหมด"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP3"
        AppendMenu hMenu, MF_STRING, 7, "ขยายทั้งหมด"
        AppendMenu hMenu, MF_STRING, 8, "ย่อลงทั้งหมด"
        'If LCase(CurrentUser.LoginName) = "admin" Then AppendMenu hMenu, MF_STRING, 4, "อนุมัติจ่าย"
        GetCursorPos Pont
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pont.x, Pont.Y, Me.hWnd, ByVal 0&)
        DestroyMenu hMenu
        
        Dim r&
        With vsMenifestList
            If ret = 1 Then
                    CopyObject = CopyData(.Row, .Col, .RowSel, .ColSel)
            ElseIf ret = 2 Then
                PastData .Row, .Col, .RowSel, .ColSel, CopyObject
            ElseIf ret = 3 Then
                SelectCheckBox 1
            ElseIf ret = 4 Then
                SelectCheckBox 2
            ElseIf ret = 5 Then
                ClosedAll = 0
                CloseMF .Row, .Col
            ElseIf ret = 6 Then
                ClosedAll = 0
                For r = 1 To .Rows - 1
                    .Row = r: .Col = 0
                    If .CellChecked = flexChecked Then CloseMF r, 1
                    .TopRow = r
                Next r
                If ClosedAll = 0 Then
                    frmMain.Stb1.Panels(1).Text = "Waste Managetment System"
                    frmMain.Stb1.Panels(2).Text = "ไม่มี Manifest ที่ถูกเลือก"
                End If
            ElseIf ret = 7 Then
                SetExtendGrid flexOutlineExpanded
            ElseIf ret = 8 Then
                SetExtendGrid flexOutlineSubtotals
            End If
        End With
    End If
End Sub

Private Sub vsMenifestList_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim r&, c&
    r = vsMenifestList.MouseRow
    c = vsMenifestList.MouseCol
    If r > 0 Then
        If c = vsMenifestList.Cols - 1 Then
            vsMenifestList.ToolTipText = "คลิกปุ่ม ... เพื่อปิดใบเมนิเฟส"
        Else
            vsMenifestList.ToolTipText = ""
        End If
    Else
        vsMenifestList.ToolTipText = ""
    End If
End Sub

Private Function CopyData(ByVal Row1 As Integer, Col1 As Integer, ByVal Row2 As Integer, ByVal Col2 As Integer) As String
    Dim CopyText$, r#, c#
    CopyText = ""
    With vsMenifestList
        For r = Row1 To Row2
            CopyColumns = 0
            For c = Col1 To Col2
                CopyColumns = CopyColumns + 1
                If CopyText = "" Then
                    CopyText = .TextMatrix(r, c)
                Else
                    CopyText = CopyText & vbTab & .TextMatrix(r, c)
                End If
            Next c
            CopyText = CopyText & vbCrLf
        Next r
        CopyText = Replace(CopyText, vbCrLf, "")
    End With
    CopyData = CopyText
End Function

Private Sub PastData(ByVal Row1 As Integer, Col1 As Integer, ByVal Row2 As Integer, ByVal Col2 As Integer, ByVal CopyText As String)
    Dim r#, c#, Columns() As String, Index As Integer
    If CopyText = "" Then Exit Sub
    Columns() = Split(CopyText, vbTab, 2)
    Index = 0
    With vsMenifestList
        For r = Row1 To Row2
            Index = 0
            For c = Col1 To Col1 + CopyColumns - 1
                If .TextMatrix(r, 1) = "H" And UCase(.TextMatrix(r, .Cols - 1)) = "PRINTED" Then
                    .TextMatrix(r, c) = Columns(Index)
                End If
                Index = Index + 1
            Next c
        Next r
    End With
End Sub

Private Sub CloseMF(ByVal Row As Long, ByVal Col As Long)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If UCase(vsMenifestList.TextMatrix(Row, vsMenifestList.Cols - 1)) <> "PRINTED" Then
        MsgBox "ไม่สามารถปิดใบเมนิเฟใบนี้ได้ กรุณาลองใหม่"
        Exit Sub
    End If
    If Trim(vsMenifestList.TextMatrix(Row, 2)) = "" Then
        MsgBox "ไม่พบเลขเมนิเฟส กรุณาลองใหม่"
        Exit Sub
    Else
        Set tmpCol = tmpS.MenifestSearch("MenifestID = '" & vsMenifestList.TextMatrix(Row, 2) & "'")
        If tmpCol.Count > 0 Then
            frmMain.Stb1.Panels(2).Text = ""
            CloseMenifest tmpCol(1), Row, Col
        Else
            frmMain.Stb1.Panels(2).Text = "ไม่พบใบเมนิเฟส..."
        End If
    End If
End Sub

Private Sub SelectCheckBox(ByVal SelectType As Byte)
    Dim i#
    For i = 1 To vsMenifestList.Rows - 1
        If vsMenifestList.TextMatrix(i, vsMenifestList.Cols - 1) = "Printed" Then
'            vsMenifestList.TextMatrix(i, 0) = SelectType
            vsMenifestList.Row = i
            vsMenifestList.Col = 0
            vsMenifestList.CellChecked = SelectType
        End If
    Next i
End Sub

Private Sub SetExtendGrid(ByVal IsCollapsed As Byte)
    Dim i#
    For i = 1 To vsMenifestList.Rows - 1
        vsMenifestList.IsCollapsed(i) = IsCollapsed
    Next i
End Sub

Private Sub SetDefaults(fa As VSFlexGrid)
    DoEvents
    With fa
        .Clear
        .Rows = 1
        .Cols = 13
        .FixedRows = 1
        .FixedCols = 0
        .BindToArray Null
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
        
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 2000
        .ColWidth(5) = 2500
        .ColWidth(6) = 1500
        .ColWidth(7) = 1100
        .ColWidth(8) = 1100
        .ColWidth(9) = 1100
        .ColWidth(10) = 1300
        .ColWidth(11) = 1500
        .ColWidth(12) = 1000
        
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "@"
        .TextMatrix(0, 2) = "ManifestID"
        .TextMatrix(0, 3) = "JobDetID"
        .TextMatrix(0, 4) = "ManifestNo"
        .TextMatrix(0, 5) = "Customer"
        .TextMatrix(0, 6) = "คิดน้ำหนักตาม"
        .TextMatrix(0, 7) = "ปริมาณผู้ก่อ"
        .TextMatrix(0, 8) = "ปริมาณผู้รับ"
        .TextMatrix(0, 9) = "วันที่ถึง Site"
        .TextMatrix(0, 10) = "เวลาที่ถึง Site"
        .TextMatrix(0, 11) = "หมายเหตุปิดงาน"
        .TextMatrix(0, 12) = "สถานะ"
        
        .ColDataType(0) = flexDTString
        .ColDataType(1) = flexDTString
        .ColDataType(2) = flexDTString
        .ColDataType(3) = flexDTString
        .ColDataType(4) = flexDTString
        .ColDataType(5) = flexDTString
        .ColDataType(6) = flexDTString
        .ColDataType(7) = flexDTString
        .ColDataType(8) = flexDTString
        .ColDataType(9) = flexDTDate
        .ColDataType(10) = flexDTDate
        .ColDataType(11) = flexDTString
        .ColDataType(12) = flexDTString
        
        .ColComboList(6) = slComboChargType
        .ColComboList(12) = "..." 'Popup for Close operator
        
        .ColFormat(9) = "##/##/####"
        .ColFormat(10) = "HH:MM"
        
        .ColEditMask(9) = "00/00/0000"
        .ColEditMask(10) = "00:00"
        .ColAlignment(9) = flexAlignRightBottom
        .ColAlignment(10) = flexAlignRightBottom
        
        ' set the properties we want
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
        .SelectionMode = flexSelectionFree
    End With
End Sub

Private Sub CloseMenifest(ByVal selMenifest As BWGMenifest, ByVal Row As Long, ByVal Col As Long)
Dim tmpStr As String, HHMM() As String, ErrMsg As String
    tmpStr = CurrentUser.VerifyAccessRole(frmMenifestClose.Name, UpdateRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr
        Exit Sub
    End If
    If vsMenifestList.TextMatrix(Row, 10) = "" Then
        frmMain.Stb1.Panels(1).Text = "ไม่ได้ระบุเวลาปิดใบเมนิเฟส ไม่สามารถปิดใบเมนิเฟสได้"
        DoEvents
        Exit Sub
    End If
    
    '*** บันทีกน้ำหนักใบเมนิเฟส
    If Trim("" & selMenifest.SignName1) <> "" Then
        selMenifest.SignName3 = selMenifest.SignName1
    Else
        selMenifest.SignName1 = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        selMenifest.SignName3 = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    End If
    If vsMenifestList.TextMatrix(Row, 6) = "ปริมาณผู้ก่อกำเนิด" Then
        selMenifest.ChargeWType = "B"
        selMenifest.ActualQty = Format(IIf(IsNumeric(vsMenifestList.TextMatrix(Row, 7)), vsMenifestList.TextMatrix(Row, 7), 0), "0.000")
        
    ElseIf vsMenifestList.TextMatrix(Row, 6) = "ปริมาณผู้รับกำจัด" Then
        selMenifest.ChargeWType = "A"
        selMenifest.ActualQty = Format(IIf(IsNumeric(vsMenifestList.TextMatrix(Row, 8)), vsMenifestList.TextMatrix(Row, 8), 0), "0.000")
    End If
    
    SetWeight selMenifest, Row, Col
    selMenifest.Save
    
    '*** บันทึกการปิดในเมนิเฟส
    HHMM() = Split(vsMenifestList.TextMatrix(Row, 10), ":")
    Load frmMenifestClose
    With frmMenifestClose
        .SetMenifestData selMenifest
        If IsDate(vsMenifestList.TextMatrix(Row, 9)) Then
            .dtDate3.ValueDMY = vsMenifestList.TextMatrix(Row, 9)
        Else
            .dtDate3.ValueDMY = FormatYMD_to_DMY(Format(Now, "dd/MM/yyyy"))
        End If
        .txtH4.Text = HHMM(0)
        .txtM4.Text = HHMM(1)
        .txtCloseDate.Text = FormatYMD_to_DMY(Format(Now, "dd/MM/yyyy"))
        ErrMsg = .SaveData(False)
        
        vsMenifestList.TextMatrix(Row, vsMenifestList.Cols - 1) = "Closed"
        vsMenifestList.Cell(flexcpChecked, Row, 0) = flexUnchecked
        SetRowStyle "Closed", Row
    End With
    ClosedAll = ClosedAll + 1
    If Trim(ErrMsg) <> "" Then
        frmMain.Stb1.Panels(1).Text = "Error " & ErrMsg
    Else
        frmMain.Stb1.Panels(1).Text = "ปิดใบเมนิเฟสเลขที่ " & selMenifest.docNo & " (ปิดแล้ว)"
    End If
    frmMain.Stb1.Panels(2).Text = "ปิดใบเมนิเฟสแล้ว " & ClosedAll & " ใบ"
    DoEvents
End Sub

Private Sub SetWeight(ByRef curMenifest_ As BWGMenifest, ByVal Row As Long, ByVal Col As Long)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim tmpKey As String, tmpVal As String
Dim tmpDet As Object
    
    If curMenifest_ Is Nothing Then
        Set tmpCol = tmpS.MenifestSearch("MenifestID = '" & vsMenifestList.TextMatrix(Row, 2) & "'")
        If tmpCol.Count > 0 Then
            Set curMenifest_ = tmpCol(1)
        Else
            frmMain.Stb1.Panels(2).Text = "ไม่พบใบเมนิเฟส..."
        End If
    End If
    
    Dim i&, hRow&, fRow&, lRow&
    Dim TotalPW As Double, TotalCW As Double
    With vsMenifestList
        For i = Row To 1 Step -1
            If .TextMatrix(i, 1) = "H" Then
                hRow = i
                fRow = i + 1
                Exit For
            End If
        Next i
        
        For i = fRow To .Rows - 1
            If i = .Rows - 1 And .TextMatrix(i, 1) = "D" Then
                lRow = .Rows - 1
                Exit For
            End If
            If .TextMatrix(i, 1) = "H" Then
                lRow = i - 1
                Exit For
            End If
        Next i
        TotalPW = 0
        TotalCW = 0
        For i = fRow To lRow
            TotalCW = TotalCW + IIf(IsNumeric(.TextMatrix(i, 7)), CDbl(.TextMatrix(i, 7)), 0)
            .TextMatrix(hRow, 7) = Format(TotalCW, "0.000")
                
            TotalPW = TotalPW + IIf(IsNumeric(Replace(.TextMatrix(i, 8), "..", ".")), CDbl(Replace(.TextMatrix(i, 8), "..", ".")), 0)
            .TextMatrix(hRow, 8) = Format(TotalPW, "0.000")
            
            tmpKey = "'" & vsMenifestList.TextMatrix(i, 3) & "'"
            If isExist(curMenifest_.JobDetails, tmpKey) Then
                Set tmpDet = curMenifest_.JobDetails(tmpKey)
            ElseIf isExist(curMenifest_.AddedWastes, tmpKey) Then
                Set tmpDet = curMenifest_.AddedWastes(tmpKey)
            End If
                    
            If Not tmpDet Is Nothing Then
                tmpVal = vsMenifestList.TextMatrix(i, 7) '*** น้ำหนักลูกค้า
                If IsNumeric(tmpVal) Then
                    tmpDet.EstWasteQty = tmpVal
                Else
                    tmpDet.EstWasteQty = 0
                End If
            End If
            
            If Not tmpDet Is Nothing Then
                tmpVal = vsMenifestList.TextMatrix(i, 8) '*** น้ำหนักบริษัท
                If IsNumeric(tmpVal) Then
                    tmpDet.DisposerWeight = tmpVal
                Else
                    tmpDet.DisposerWeight = 0
                End If
            End If
            
        Next i
    End With
End Sub
