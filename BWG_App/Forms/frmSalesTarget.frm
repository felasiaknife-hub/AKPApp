VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesTarget 
   Caption         =   "เป้าการขาย"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   11025
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   9000
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   15875
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VSFlex7LCtl.VSFlexGrid VSGSaleTarget 
      Height          =   9000
      Left            =   3045
      TabIndex        =   0
      Top             =   30
      Width           =   7950
      _cx             =   14023
      _cy             =   15875
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
End
Attribute VB_Name = "frmSalesTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colTarget_ As Collection
Dim slCompanyCom As String
Const InPro As String = "|#0*1;ผ่านโปร|#1;ทดลองงาน"
Const IsOverall As String = "|#0*1;ยอดเฉพาะ|#1;ยอดรวม"
Dim AddNewRow As Integer



Private Sub Form_Load()
    PopulateCompany
    InitializeData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    SSSplitter1.Height = Me.Height - frmMain.Stb1.Height
'    SSSplitter1.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage True
End Sub

Private Sub PopulateCompany()
Dim X%
    slCompanyCom = ""
    For X = 1 To CurrentUser.WorkCompanies.Count
        If X = 1 Then
            slCompanyCom = "#" & CurrentUser.WorkCompanies(X).ID & "*" & X & ";" & CurrentUser.WorkCompanies(X).CompanyNameTH
        Else
            slCompanyCom = slCompanyCom & "|#" & CurrentUser.WorkCompanies(X).ID & ";" & CurrentUser.WorkCompanies(X).CompanyNameTH
        End If
    Next
End Sub

Private Sub InitializeData()
    On Error GoTo ErrD
    Dim RS As ADODB.Recordset
    Dim tSql As String, i As Integer
    Dim xNode As MSComctlLib.Node
    Dim xParent As String, PreviousNd As String
    TreeView1.Nodes.Clear
    Set xNode = TreeView1.Nodes.Add(, , "A", "Commission Setting...")
    xNode.Expanded = True
    tSql = "SELECT * FROM vw_SaleDistrictSearch ORDER BY ZoneName, SDName"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    Do While Not RS.EOF
        xParent = RS!ZoneName
        If RS.AbsolutePosition = 1 Then
            Set xNode = TreeView1.Nodes.Add("A", tvwChild, RS!ZoneName, RS!ZoneName)
            Set xNode = TreeView1.Nodes.Add(xParent, tvwChild, RS!SaleStaffID & "_" & RS!SaleDistrictID, RS!SDName & " [" & RS!StaffFName & " " & RS!StaffLName & "]")
            PreviousNd = RS!ZoneName
        Else
            If PreviousNd = RS!ZoneName Then
                Set xNode = TreeView1.Nodes.Add(xParent, tvwChild, RS!SaleStaffID & "_" & RS!SaleDistrictID, RS!SDName & " [" & RS!StaffFName & " " & RS!StaffLName & "]")
            Else
                Set xNode = TreeView1.Nodes.Add("A", tvwChild, RS!ZoneName, RS!ZoneName)
                Set xNode = TreeView1.Nodes.Add(xParent, tvwChild, RS!SaleStaffID & "_" & RS!SaleDistrictID, RS!SDName & " [" & RS!StaffFName & " " & RS!StaffLName & "]")
            End If
            PreviousNd = RS!ZoneName
        End If
        RS.MoveNext
    Loop
    Set RS = Nothing
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrHandlers
    Dim tmpS As New BWGSearchManager
    Dim tmpTarget As BWGSaleTarget
    Dim RS As ADODB.Recordset, X As Integer
    Dim tSql As String, SaleStaffID As String, SaleDistrictID As String
    Dim slTargetID As String, slCompany As String, slForYear As String
    Dim slForMonth As String, slZoneName As String, slSDName As String, slSaleName As String
    Dim slNonTarget As String, slNBLTarget As String, slHazTarget As String, slHBLTarget As String
    Dim slHICTarget As String, slHBITarget As String, slBBHTarget As String, slBBLTarget As String
    Dim slInPro As String, slIsOverall As String, slTotal As String
    If Node.Key = "A" Then Exit Sub
    If Node.Parent.Key = "A" Then Exit Sub
    
    SaleStaffID = Replace(Trim(Split(Node.Key, "_")(0)), "'", "")
    SaleDistrictID = Replace(Trim(Split(Node.Key, "_")(1)), "'", "")
    
    Set colTarget_ = tmpS.SaleTargetSearch("SaleStaffID = '" & SaleStaffID & "' AND SaleDistrictID = '" & SaleDistrictID & "'")
    If colTarget_.Count = 0 Then
        SetDefaults VSGSaleTarget
        MsgBox "ไม่มีข้อมูล", vbInformation
        Exit Sub
    End If
    For X = 1 To colTarget_.Count
        Set tmpTarget = colTarget_(X)
        slTargetID = slTargetID & "'" & tmpTarget.ID & "'" & "|"
        slCompany = slCompany & tmpTarget.CompanyName & "|"
        slForYear = slForYear & tmpTarget.ForYear & "|"
        slZoneName = slZoneName & tmpTarget.SaleZoneName & "|"
        slSDName = slSDName & tmpTarget.SaleDistrictName & "|"
        slSaleName = slSaleName & tmpTarget.SaleStaffName & "|"
        slForMonth = slForMonth & tmpTarget.ForMonth & "|"
        slNonTarget = slNonTarget & tmpTarget.NonTarget & "|"
        slNBLTarget = slNBLTarget & tmpTarget.NBLTarget & "|"
        slHazTarget = slHazTarget & tmpTarget.HzTarget & "|"
        slHBLTarget = slHBLTarget & tmpTarget.HBLTarget & "|"
        slHICTarget = slHICTarget & tmpTarget.HICTarget & "|"
        slHBITarget = slHBITarget & tmpTarget.HBITarget & "|"
        slBBHTarget = slBBHTarget & tmpTarget.BBHTarget & "|"
        slBBLTarget = slBBLTarget & tmpTarget.BBLTarget & "|"
        slTotal = slTotal & (tmpTarget.NonTarget + tmpTarget.NBLTarget + tmpTarget.HzTarget + tmpTarget.HBLTarget + tmpTarget.HICTarget + tmpTarget.HBITarget + tmpTarget.BBHTarget + tmpTarget.BBLTarget) & "|"
        slInPro = slInPro & tmpTarget.InPro & "|"
        slIsOverall = slIsOverall & tmpTarget.IsOverall & "|"
        Set tmpTarget = Nothing
    Next X
    
    slTargetID = "TargetID|" & IIf(Right(slTargetID, 1) = "|", Left(slTargetID, Len(slTargetID) - 1), slTargetID)
    slCompany = "บริษัท|" & IIf(Right(slCompany, 1) = "|", Left(slCompany, Len(slCompany) - 1), slCompany)
    slForYear = "ปี|" & IIf(Right(slForYear, 1) = "|", Left(slForYear, Len(slForYear) - 1), slForYear)
    slZoneName = "โซน|" & IIf(Right(slZoneName, 1) = "|", Left(slZoneName, Len(slZoneName) - 1), slZoneName)
    slSDName = "เขต|" & IIf(Right(slSDName, 1) = "|", Left(slSDName, Len(slSDName) - 1), slSDName)
    slSaleName = "พนักงานขาย|" & IIf(Right(slSaleName, 1) = "|", Left(slSaleName, Len(slSaleName) - 1), slSaleName)
    slForMonth = "เดือน|" & IIf(Right(slForMonth, 1) = "|", Left(slForMonth, Len(slForMonth) - 1), slForMonth)
    slNonTarget = "Non|" & IIf(Right(slNonTarget, 1) = "|", Left(slNonTarget, Len(slNonTarget) - 1), slNonTarget)
    slHazTarget = "Haz|" & IIf(Right(slHazTarget, 1) = "|", Left(slHazTarget, Len(slHazTarget) - 1), slHazTarget)
    slNBLTarget = "NBL|" & IIf(Right(slNBLTarget, 1) = "|", Left(slNBLTarget, Len(slNBLTarget) - 1), slNBLTarget)
    slHBLTarget = "HBL|" & IIf(Right(slHBLTarget, 1) = "|", Left(slHBLTarget, Len(slHBLTarget) - 1), slHBLTarget)
    slHICTarget = "HIC|" & IIf(Right(slHICTarget, 1) = "|", Left(slHICTarget, Len(slHICTarget) - 1), slHICTarget)
    slHBITarget = "HBI|" & IIf(Right(slHBITarget, 1) = "|", Left(slHBITarget, Len(slHBITarget) - 1), slHBITarget)
    slBBHTarget = "BBH|" & IIf(Right(slBBHTarget, 1) = "|", Left(slBBHTarget, Len(slBBHTarget) - 1), slBBHTarget)
    slBBLTarget = "BBL|" & IIf(Right(slBBLTarget, 1) = "|", Left(slBBLTarget, Len(slBBLTarget) - 1), slBBLTarget)
    slTotal = "TOTAL|" & IIf(Right(slTotal, 1) = "|", Left(slTotal, Len(slTotal) - 1), slTotal)
    slInPro = "Pro|" & IIf(Right(slInPro, 1) = "|", Left(slInPro, Len(slInPro) - 1), slInPro)
    slIsOverall = "OverAll|" & IIf(Right(slIsOverall, 1) = "|", Left(slIsOverall, Len(slIsOverall) - 1), slIsOverall)
    
    ' reset the control
    SetDefaults VSGSaleTarget
    
    ' set the properties we want
    VSGSaleTarget.Rows = colTarget_.Count + 1
    VSGSaleTarget.Cols = 18
    VSGSaleTarget.FixedRows = 1
'    VSGSaleTarget.FixedCols = 0
    VSGSaleTarget.AllowUserResizing = flexResizeBoth
    VSGSaleTarget.ExtendLastCol = True
    VSGSaleTarget.ExplorerBar = flexExMove ' can't sort when bound
    VSGSaleTarget.Editable = True
    
    'set column properties
    VSGSaleTarget.ColComboList(1) = slCompanyCom
    VSGSaleTarget.ColComboList(16) = InPro
    VSGSaleTarget.ColComboList(17) = IsOverall
    
    ' add data
    FillColumn VSGSaleTarget, 0, slTargetID
    FillColumn VSGSaleTarget, 1, slCompany
    FillColumn VSGSaleTarget, 2, slForYear
    FillColumn VSGSaleTarget, 3, slZoneName
    FillColumn VSGSaleTarget, 4, slSDName
    FillColumn VSGSaleTarget, 5, slSaleName
    FillColumn VSGSaleTarget, 6, slForMonth
    FillColumn VSGSaleTarget, 7, slNonTarget
    FillColumn VSGSaleTarget, 8, slHazTarget
    FillColumn VSGSaleTarget, 9, slNBLTarget
    FillColumn VSGSaleTarget, 10, slHBLTarget
    FillColumn VSGSaleTarget, 11, slHICTarget
    FillColumn VSGSaleTarget, 12, slHBITarget
    FillColumn VSGSaleTarget, 13, slBBHTarget
    FillColumn VSGSaleTarget, 14, slBBLTarget
    FillColumn VSGSaleTarget, 15, slTotal
    FillColumn VSGSaleTarget, 16, slInPro
    FillColumn VSGSaleTarget, 17, slIsOverall
    VSGSaleTarget.ColFormat(7) = "#,###"
    VSGSaleTarget.ColFormat(8) = "#,###"
    VSGSaleTarget.ColFormat(9) = "#,###"
    VSGSaleTarget.ColFormat(10) = "#,###"
    VSGSaleTarget.ColFormat(11) = "#,###"
    VSGSaleTarget.ColFormat(12) = "#,###"
    VSGSaleTarget.ColFormat(13) = "#,###"
    VSGSaleTarget.ColFormat(14) = "#,###"
    VSGSaleTarget.ColFormat(15) = "#,###"
    
    ' sort and total
    SubtotalsSort
    
    Set RS = Nothing
    Exit Sub
ErrHandlers:
    MsgBox err.Description
    Resume Next
End Sub

Sub SubtotalsSort()
    ' sort
'    VSGSaleTarget.Select 1, 0, 1, VSGSaleTarget.Cols - 1
    VSGSaleTarget.Sort = flexSortCustom
    VSGSaleTarget.Select 0, 0, 0, 0
    
    ' insert subtotals
    VSGSaleTarget.OutlineCol = 0
    VSGSaleTarget.OutlineBar = flexOutlineBarComplete
    VSGSaleTarget.SubtotalPosition = flexSTAbove
    VSGSaleTarget.Subtotal flexSTClear
'    VSGSaleTarget.Subtotal flexSTSum, -1, 6, "#,", 1, vbWhite, True
'    VSGSaleTarget.Subtotal flexSTSum, 0, 6, "#,", &H404040, vbWhite, True
'    VSGSaleTarget.Subtotal flexSTSum, 1, 6, "#,", &H808080, vbWhite, False
'    VSGSaleTarget.Subtotal flexSTSum, 2, 6, "#,", &HC0C0C0, vbWhite, False
'    VSGSaleTarget.Subtotal flexSTSum, 3, 6, "#,", &HC0C0C0, vbWhite, False
'    VSGSaleTarget.Subtotal flexSTSum, 4, 6, "#,", &HC0C0C0, vbWhite, False
'    VSGSaleTarget.Subtotal flexSTSum, 5, 6, "#,", &HC0C0C0, vbWhite, False
    
    ' auto size
    VSGSaleTarget.ColWidth(0) = 0
    VSGSaleTarget.AutoSize 1, VSGSaleTarget.Cols - 1
    
    ' merge
    VSGSaleTarget.MergeCells = flexMergeRestrictAll
'    VSGSaleTarget.MergeCol(0) = True
'    VSGSaleTarget.MergeCol(1) = True
'    VSGSaleTarget.MergeCol(2) = True
'    VSGSaleTarget.MergeCol(3) = True
'    VSGSaleTarget.MergeCol(4) = True
'    VSGSaleTarget.MergeCol(5) = True
'    VSGSaleTarget.MergeCol(6) = True
End Sub

Sub SetDefaults(fa As VSFlexGrid)
    With fa
        .BindToArray Null
        .Rows = 0
        .Cols = 0
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        .Editable = False
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
    End With
End Sub

Sub FillColumn(fa As VSFlexGrid, ByVal c As Long, ByVal s As String)
    Dim r&, i&, cnt&
    ReDim lst(0) As String
    
    ' build list
    cnt = 0
    i = InStr(s, "|")
    While i > 0
        lst(cnt) = Left$(s, i - 1)
        s = Mid$(s, i + 1)
        cnt = cnt + 1
        ReDim Preserve lst(cnt) As String
        i = InStr(s, "|")
    Wend
    lst(cnt) = s
    
    ' set values
    fa.Cell(flexcpText, 0, c) = lst(0)
    For r = fa.FixedRows To fa.Rows - 1
        i = r '(Rnd() * 1000) Mod cnt + 1
        fa.Cell(flexcpText, r, c) = lst(i)
    Next
    
    ' do an autosize
    fa.AutoSize c, , , 300
End Sub



Private Sub VSGSaleTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim r&, c&
    If Button <> vbRightButton Then Exit Sub
    If VSGSaleTarget.MouseRow = r Then Exit Sub
    If VSGSaleTarget.MouseRow = -1 Then Exit Sub
    r = VSGSaleTarget.MouseRow
    VSGSaleTarget.Select r, 0, r, VSGSaleTarget.Cols - 1
End Sub

Private Sub VSGSaleTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim Pt As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "Add Target"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP1"
        AppendMenu hMenu, MF_STRING, 2, "Delete Target"
'        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP2"
'        AppendMenu hMenu, MF_STRING, 3, "ไม่อนุมัติจ่าย"
'        If LCase(CurrentUser.LoginName) = "admin" Then AppendMenu hMenu, MF_STRING, 4, "อนุมัติจ่าย"
        GetCursorPos Pt
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pt.X, Pt.y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu
        
        If ret = 1 Then
            AddTarget
        ElseIf ret = 2 Then
            DelTarget
        End If
    End If
End Sub

Private Sub AddTarget()
    With VSGSaleTarget
        .AddItem .TextMatrix(.Rows - 1, 0) & vbTab & .TextMatrix(.Rows - 1, 1) & vbTab & .TextMatrix(.Rows - 1, 2) & vbTab & .TextMatrix(.Rows - 1, 3) & vbTab & .TextMatrix(.Rows - 1, 4), VSGSaleTarget.Rows
        AddNewRow = .Rows
    End With
End Sub

Private Sub DelTarget()
    Dim tmpTarget As BWGSaleTarget
    Dim tmpKey As String
    If MsgBox("ลบข้อมูลเป้าการขาย", vbInformation + vbYesNo) = vbNo Then Exit Sub
    With VSGSaleTarget
        tmpKey = .TextMatrix(.RowSel, 0)
        If isExist(colTarget_, tmpKey) = False Then MsgBox "ไม่พบข้อมูล", vbInformation: Exit Sub
        Set tmpTarget = colTarget_(tmpKey)
        tmpTarget.Delete
        .RemoveItem .RowSel
    End With
End Sub
