VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmBoxManaulGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "บันทึกหมายเลข BOX"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   9075
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10995
      Begin VB.CommandButton cmdAdd 
         Caption         =   "เพิ่ม Box"
         Height          =   435
         Left            =   6180
         TabIndex        =   22
         Top             =   900
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Gen"
         Height          =   255
         Left            =   7380
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Height          =   1515
         Left            =   7320
         TabIndex        =   17
         Top             =   240
         Width           =   3375
         Begin VB.TextBox txtAddFrom 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "รันต่อจากหมายเลข >>"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   540
            Width           =   1875
         End
         Begin VB.TextBox txtCount 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   1140
            Width           =   675
         End
         Begin VB.TextBox txtFormat 
            Height          =   315
            Left            =   840
            TabIndex        =   8
            Top             =   1140
            Width           =   675
         End
         Begin VB.CommandButton cmdExcute 
            Caption         =   "ดำเนินการ"
            Height          =   495
            Left            =   2280
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtPreIndex 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1140
            Width           =   675
         End
         Begin VB.OptionButton Option1 
            Caption         =   "รันต่อจากหมายเลขล่าสุด"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.Label lblCount 
            Caption         =   "จำนวน"
            Height          =   195
            Left            =   1560
            TabIndex        =   20
            Top             =   900
            Width           =   675
         End
         Begin VB.Label lblFormat 
            Caption         =   "ฟอร์แมต"
            Height          =   195
            Left            =   840
            TabIndex        =   19
            Top             =   900
            Width           =   675
         End
         Begin VB.Label lblPre 
            Caption         =   "PreIndex"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   900
            Width           =   675
         End
      End
      Begin VB.ComboBox cboBoxOwner 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   240
         Width           =   5955
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   240
         TabIndex        =   13
         Top             =   8280
         Width           =   10455
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   9180
            TabIndex        =   24
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   7860
            TabIndex        =   23
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6615
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   10455
         Begin VSFlex7LCtl.VSFlexGrid vsgBox 
            Height          =   6375
            Left            =   60
            TabIndex        =   21
            Top             =   180
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
      End
      Begin VB.TextBox txtBoxNo 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   1635
      End
      Begin VB.ComboBox cboBoxType 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "หมายเลขบ็อก"
         Height          =   195
         Left            =   4500
         TabIndex        =   16
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label4 
         Caption         =   "ประเภทบ็อก"
         Height          =   195
         Left            =   1260
         TabIndex        =   15
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "ชื่อเจ้าของบ็อก"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBoxManaulGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TypeOfBox As String
Const StateOfBox As String = "#0;ถูกจัดไปแล้ว|#1;พร้อมใช้งาน|#2;ชำรุด|#3;หาย|#4;จำหน่าย"

Dim colSearchResult_ As Collection

Private Sub cboBoxOwner_Click()
    ValidateCbo
End Sub

Private Sub ValidateCbo()
    Dim BoxItem As New BWGBoxDataItem
    Dim RS As ADODB.Recordset
    Dim tSql As String, OrderBy As String, x As Integer
    If cboBoxOwner.ListIndex <= 0 Then Exit Sub
    vsgBox.Rows = 2
    tSql = "BoxOwnersCode = '" & cboBoxOwner.List(cboBoxOwner.ListIndex) & "'"
    If cboBoxType.ListIndex >= 0 Then tSql = tSql & " AND BoxType = '" & cboBoxType.List(cboBoxType.ListIndex) & "'"
    OrderBy = "ORDER BY BoxType, Len(BoxCode) DESC, BoxCode ASC"
    Set colSearchResult_ = BoxItem.BoxDataItemSearch(tSql, OrderBy)
'    vsgBox.Rows = colSearchResult_.Count
    For x = 1 To colSearchResult_.Count
        Set BoxItem = colSearchResult_(x)
        With vsgBox
            .TextMatrix(x, 0) = BoxItem.BoxOwnersCode
            .TextMatrix(x, 1) = BoxItem.BoxType
            .TextMatrix(x, 2) = BoxItem.BoxNo
            .TextMatrix(x, 3) = BoxItem.BoxStatus
            If .Rows - 1 = x Then .Rows = .Rows + 1
        End With
        Set BoxItem = Nothing
    Next x
    ClearText
End Sub

Private Sub ClearText()
    Check1.Value = 0
    txtPreIndex.Text = ""
    txtFormat.Text = ""
    txtCount.Text = ""
End Sub

Private Sub cboBoxOwner_Validate(Cancel As Boolean)
    ValidateCbo
End Sub

Private Sub cboBoxType_Click()
    ValidateCbo
End Sub

Private Sub cboBoxType_Validate(Cancel As Boolean)
    ValidateCbo
End Sub

Private Sub Check1_Click()
    Frame4.Enabled = CBool(Check1)
    Option1.Value = CBool(Check1)
    txtAddFrom.Text = ""
End Sub

Private Sub cmdAdd_Click()
    If cboBoxOwner.ListIndex = 0 Then
        MsgBox "กรุณาเลือกเจ้าของบ็อก", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtBoxNo.Text) = "" Then
        MsgBox "กรุณาระบุหมยเลขบ็อก", vbInformation
        Exit Sub
    End If
    AddToGrid
    txtBoxNo.Text = ""
    txtBoxNo.SetFocus
End Sub

Private Sub cmdExcute_Click()
    'Check & Get Last BoxOwners BoxType
    On Error Resume Next
    If cboBoxOwner.ListIndex <= 0 Then
        MsgBox "Please select Box Ownser.", vbInformation
        Exit Sub
    End If
    If cboBoxType.ListIndex <= 0 Then
        MsgBox "Please select BoxType", vbInformation
        Exit Sub
    End If
    
    If Option2.Value = True Then
        If txtAddFrom = "" Then
            MsgBox "ใส่หมายเลขที่ต้องการรันต่อจากหมายเลขที่มีในระบบ หรือ เริ่มต้นจาก 0 ในกรณีที่ Format เลข Box ไม่เท่ากัน", vbInformation
            Exit Sub
        End If
    End If
    
    Dim RS As ADODB.Recordset
    Dim tSql As String, MaxBoxCode As String, NewRunNo As Integer
    Dim XX As Integer, StrFormat As String
    tSql = "SELECT MAX(BoxCode) AS MaxBoxCode FROM Box_BoxNo WHERE BoxOwnersCode = '" & cboBoxOwner.Text & "' AND BoxType = '" & cboBoxType.Text & "'"
    If Option2.Value = True Then
        tSql = tSql & vbCrLf & "AND BoxCode ='" & Trim(txtAddFrom) & "'"
    End If
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    If RS.EOF Then
        If Option2.Value = True Then
            MaxBoxCode = Trim(txtAddFrom)
        Else
            MaxBoxCode = ""
        End If
    Else
        If Option2.Value = True Then
            MaxBoxCode = Trim(txtAddFrom)
        Else
            MaxBoxCode = RS!MaxBoxCode
        End If
    End If
    
    StrFormat = ""
    NewRunNo = 0
    For XX = 1 To Len(txtFormat)
        StrFormat = StrFormat & "0"
    Next
    If MaxBoxCode <> "" Then
        For XX = 1 To Len(MaxBoxCode)
            MaxBoxCode = Right(MaxBoxCode, Len(MaxBoxCode) - XX)
            If IsNumeric(MaxBoxCode) Then
                If CInt(MaxBoxCode) >= 0 Then
                    NewRunNo = CDbl(MaxBoxCode)
                    Exit For
                End If
            End If
        Next XX
    Else
        NewRunNo = 0
    End If
    
    If Trim(txtCount) = "" Then txtCount = "0"
    For XX = 1 To txtCount
        NewRunNo = NewRunNo + 1
        MaxBoxCode = Format(NewRunNo, StrFormat)
        txtBoxNo.Text = txtPreIndex & MaxBoxCode
        vsgBox.SetFocus
        cmdAdd_Click
    Next XX
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandlers
    Dim RS As ADODB.Recordset
    Dim tSql As String
    Dim i As Integer, iCount As Integer
    Dim BoxItem As New BWGBoxDataItem
'    With vsgBox
'        For i = 1 To .Rows - 1
'            If .TextMatrix(i, 3) = "N" Then
'                tSql = "INSERT INTO Box_BoxNo(BoxID, BoxOwnersCode, BoxType, BoxCode, BoxStatus, BoxDesc)" & vbCrLf
'                tSql = tSql & "SELECT '" & .TextMatrix(i, 0) & .TextMatrix(i, 1) & .TextMatrix(i, 2) & "', '" & .TextMatrix(i, 0) & "', '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 2) & "', '1', ''"
'                GetDBConnection DBConStock
'                DBConn.Execute tSql
'                iCount = iCount + 1
'                .TextMatrix(i, 3) = "1"
'            End If
'        Next i
'    End With
    
    For i = 1 To colSearchResult_.Count
        Set BoxItem = colSearchResult_(i)
        If Not BoxItem Is Nothing Then
            If BoxItem.isChanged Then
                BoxItem.Save
                iCount = iCount + 1
            End If
        End If
    Next i
    
    MsgBox "บันทึกทั้งหมด " & iCount & "รายการ", vbInformation
    ClearText
    Exit Sub
ErrHandlers:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim RS As ADODB.Recordset
    Dim tSql As String
    vsgBox.Cols = 4
    vsgBox.Rows = 2
    vsgBox.FixedCols = 0
    vsgBox.FixedRows = 0
    vsgBox.ColWidth(0) = 1000
    vsgBox.ColWidth(1) = 1000
    vsgBox.ColWidth(2) = 2000
    vsgBox.ColWidth(3) = 2000
    vsgBox.ColAlignment(2) = 1
    vsgBox.TextMatrix(0, 0) = "เจ้าของบ็อก"
    vsgBox.TextMatrix(0, 1) = "ประเภทบ็อก"
    vsgBox.TextMatrix(0, 2) = "หมายเลขบ็อก"
    vsgBox.TextMatrix(0, 3) = "สถานะบ็อก"
    vsgBox.ColComboList(3) = StateOfBox
    vsgBox.Editable = flexEDKbdMouse
    
    tSql = "SELECT * FROM Box_BoxOwners ORDER BY BoxOwnersCode"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    cboBoxOwner.AddItem "== เลือกเจ้าของบ็อก =="
    Do While Not RS.EOF
        cboBoxOwner.AddItem RS!BoxOwnersCode
        RS.MoveNext
    Loop
    If cboBoxOwner.ListCount > 0 Then cboBoxOwner.ListIndex = 0
    
    LoadBoxType
    Check1.Value = 0
    Check1_Click
End Sub

Private Sub LoadBoxType()
    Dim RS As New ADODB.Recordset
    Dim tSql As String, i As Integer
    tSql = "SELECT * FROM Box_BoxType ORDER BY BoxType"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    Do While Not RS.EOF
        TypeOfBox = RS!BoxType
        cboBoxType.AddItem TypeOfBox
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub AddToGrid()
    Dim i As Integer, tmpKey As String
    Dim BoxItem As New BWGBoxDataItem
    With vsgBox
        For i = 1 To vsgBox.Rows - 1
            If .TextMatrix(i, 0) = "" Then
                tmpKey = cboBoxOwner.List(cboBoxOwner.ListIndex) & cboBoxType.List(cboBoxType.ListIndex) & Trim(txtBoxNo.Text)
                If isExist(colSearchResult_, tmpKey) Then
                    MsgBox "บ็อกนี้มีแล้ว", vbInformation
                    Exit Sub
                End If
                .TextMatrix(i, 0) = cboBoxOwner.List(cboBoxOwner.ListIndex)
                .TextMatrix(i, 1) = cboBoxType.List(cboBoxType.ListIndex)
                .TextMatrix(i, 2) = Trim(txtBoxNo.Text)
                .TextMatrix(i, 3) = "1"
                .Select i, 0, i, .Cols - 1
                
                BoxItem.BoxID = cboBoxOwner.List(cboBoxOwner.ListIndex) & cboBoxType.List(cboBoxType.ListIndex) & Trim(txtBoxNo.Text)
                BoxItem.BoxOwnersCode = cboBoxOwner.List(cboBoxOwner.ListIndex)
                BoxItem.BoxType = cboBoxType.List(cboBoxType.ListIndex)
                BoxItem.BoxNo = Trim(txtBoxNo.Text)
                BoxItem.BoxStatus = 1
                BoxItem.isChanged = True
                colSearchResult_.Add BoxItem
                DoEvents
                Exit For
            End If
        Next i
        If .Rows - 1 = i Then .Rows = .Rows + 1
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colSearchResult_ = Nothing
End Sub

Private Sub Option1_Click()
    txtAddFrom.Visible = False
    txtAddFrom.Text = ""
End Sub

Private Sub Option2_Click()
    txtAddFrom.Visible = True
End Sub

Private Sub txtBoxNo_GotFocus()
    cmdAdd.Default = True
End Sub

Private Sub txtBoxNo_LostFocus()
    cmdAdd.Default = False
End Sub

Private Sub vsgBox_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim BoxItem As New BWGBoxDataItem, tmpKey As String
    Dim UGroup As String, x As Integer
    If Col <> 3 Then Exit Sub
    tmpKey = "'" & vsgBox.TextMatrix(Row, 0) & vsgBox.TextMatrix(Row, 1) & vsgBox.TextMatrix(Row, 2) & "'"
    Set BoxItem = colSearchResult_(tmpKey)
    UGroup = ""
    For x = 1 To CurrentUser.getUserGroups.Count
        If CurrentUser.getUserGroups(x).Name = "Senior Transport" Then
            UGroup = CurrentUser.getUserGroups(x).Name
            Exit For
        End If
    Next x
    
    If BoxItem.BoxStatus = 0 Then
        If UGroup <> "Senior Transport" Then
            MsgBox "Box ใบนี้ ไม่สามารถเปลี่ยนสถานะได้เนื่องจากได้จัดไปให้ลูกค้าแล้ว, ถ้าต้องการเปลี่ยนสถานะให้ทำการยกเลิกการจัดบ็อก", vbInformation
            Exit Sub
        End If
    End If
    BoxItem.BoxStatus = vsgBox.TextMatrix(Row, Col)
    Set BoxItem = Nothing
End Sub

Private Sub vsgBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tSql As String, BoxID As String
    If KeyCode = vbKeyDelete Then
        If vsgBox.TextMatrix(vsgBox.RowSel, vsgBox.ColSel) = "" Then Exit Sub
        BoxID = vsgBox.TextMatrix(vsgBox.RowSel, 0) & vsgBox.TextMatrix(vsgBox.RowSel, 1) & vsgBox.TextMatrix(vsgBox.RowSel, 2)
        tSql = "DELETE FROM Box_BoxNo WHERE BoxID = '" & BoxID & "'"
        If MsgBox("ยืนยันการลบข้อมูล", vbYesNo) = vbNo Then Exit Sub
        GetDBConnection DBConStock
        DBConn.Execute tSql
        vsgBox.RemoveItem vsgBox.RowSel
    End If
End Sub
