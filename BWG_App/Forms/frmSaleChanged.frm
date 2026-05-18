VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSaleChanged 
   Caption         =   "Sale Changed Table"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkBox 
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   4980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6300
      Left            =   0
      ScaleHeight     =   6270
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update to New SD"
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   3720
         Width           =   2295
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frmSaleChanged.frx":0000
         Left            =   180
         List            =   "frmSaleChanged.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   420
         Width           =   2325
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   885
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   4635
      Left            =   2640
      TabIndex        =   0
      Top             =   -60
      Width           =   9015
      _cx             =   15901
      _cy             =   8176
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
Attribute VB_Name = "frmSaleChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaleStaff_ As BWGCompanyStaff
Dim colDistrict_ As Collection
Dim SDList As String

Private Sub SetGrid()
    With VSFlexGrid1
        .Left = Picture1.Width
        .Top = 0
        .Cols = 14
        .Rows = 10
        .FixedCols = 1
        .FixedRows = 1
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "CustomerID"
        .TextMatrix(0, 2) = "Cust Code"
        .TextMatrix(0, 3) = "Cust Name"
        .TextMatrix(0, 4) = "SaleStaffID"
        .TextMatrix(0, 5) = "SaleDistrictID"
        .TextMatrix(0, 6) = "SD Name"
        .TextMatrix(0, 7) = "Saleman"
        .TextMatrix(0, 8) = "New SaleStaffID"
        .TextMatrix(0, 9) = "New SaleDistrictID"
        .TextMatrix(0, 10) = "New SD Name"
        .TextMatrix(0, 11) = "New Saleman"
        .TextMatrix(0, 12) = "Changed By"
        .TextMatrix(0, 13) = "Changed Date"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 1100
        .ColWidth(3) = 3750
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 1100
        .ColWidth(7) = 2000
        .ColWidth(8) = 0
        .ColWidth(9) = 2000
        .ColWidth(10) = 1200
        .ColWidth(11) = 2000
        .ColWidth(12) = 1500
        .ColWidth(13) = 1200
        
        .ColAlignment(0) = flexAlignCenterTop
        .ColAlignment(6) = flexAlignLeftTop
        .ColAlignment(10) = flexAlignLeftTop
        .ColAlignment(12) = flexAlignLeftTop
        .ColAlignment(13) = flexAlignLeftTop
        
        .ColDataType(0) = flexDTBoolean
        .ColComboList(9) = SDList
        .Editable = flexEDKbdMouse
        
        .AutoResize = True
        .AllowSelection = True
        .AutoSizeMode = flexAutoSizeColWidth
                        
    End With
    VSFlexGrid1.Row = 0
    VSFlexGrid1.Col = 0
    ChkBox.Top = VSFlexGrid1.Top + VSFlexGrid1.CellTop + 50
    ChkBox.Left = VSFlexGrid1.Left + 100
    ChkBox.Height = VSFlexGrid1.CellHeight - 20
    ChkBox.Visible = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim RS As New ADODB.Recordset
    Dim tSql As String
    
    tSql = "SELECT CustomerID, CustomerCode, CompanyName, UnderSaleStaffID, SaleDistrictID, SDName = (SELECT SDName FROM tbSaleDistrict WHERE SaleDistrictID = vw_CustomerAll.SaleDistrictID), SaleName = (SELECT StaffFName + ' ' + StaffLName FROM tbCompanyStaff WHERE StaffID = vw_CustomerAll.UnderSaleStaffID), [Status], LastUser, LastUpdate" & vbCrLf
    tSql = tSql & "FROM vw_CustomerAll" & vbCrLf
    
    If Not SaleStaff_ Is Nothing Then
        tSql = tSql & "WHERE UnderSaleStaffID = '" & SaleStaff_.ID & "' " & vbCrLf
'        tSql = tSql & "AND SaleDistrictID = '" & "" & "'"
    End If
    tSql = tSql & "ORDER BY CompanyName"
    
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    With VSFlexGrid1
        .Rows = RS.RecordCount + 1
        .Clear 1, 1
        DoEvents
        Do While Not RS.EOF
            .TextMatrix(RS.AbsolutePosition, 0) = ""
            .TextMatrix(RS.AbsolutePosition, 1) = Trim("" & RS!CustomerID)
            .TextMatrix(RS.AbsolutePosition, 2) = Trim("" & RS!CustomerCode)
            .TextMatrix(RS.AbsolutePosition, 3) = Trim("" & RS!CompanyName)
            .TextMatrix(RS.AbsolutePosition, 4) = Trim("" & RS!UnderSaleStaffID)
            .TextMatrix(RS.AbsolutePosition, 5) = Trim("" & RS!SaleDistrictID)
            .TextMatrix(RS.AbsolutePosition, 6) = Trim("" & RS!SDName)
            .TextMatrix(RS.AbsolutePosition, 7) = Trim("" & RS!salename)
            .TextMatrix(RS.AbsolutePosition, 8) = ""
            .TextMatrix(RS.AbsolutePosition, 9) = ""
            .TextMatrix(RS.AbsolutePosition, 10) = ""
            .TextMatrix(RS.AbsolutePosition, 11) = ""
            .TextMatrix(RS.AbsolutePosition, 12) = Trim("" & RS!LastUser)
            .TextMatrix(RS.AbsolutePosition, 13) = Trim("" & RS!LastUpdate)
            RS.MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Private Sub cmdUpdate_Click()
    Dim tSql As String
    Dim x As Long
    If MsgBox("ยืนยันการอัพเดทข้อมูล ข้อมูลที่ถูกทำเครื่องถูกด้านหน้าจะถูกอัพเดททั้งหมด", vbInformation + vbYesNo) = vbNo Then Exit Sub
    For x = 1 To VSFlexGrid1.Rows
        If CBool(VSFlexGrid1.TextMatrix(x, 0)) = True Then
            If Trim(VSFlexGrid1.TextMatrix(x, 10)) <> "" Then
                tSql = "UPDATE tbCustomer SET UnderSaleStaffID = '" & "" & "', SaleDistrictID = '" & "" & "'" & vbCrLf
                tSql = tSql & "WHERE SaleDistrictID = '" & Trim(VSFlexGrid1.TextMatrix(x, 4)) & "' AND UnderSaleStaffID = '" & Trim(VSFlexGrid1.TextMatrix(x, 5)) & "'"
                DBConn.Execute tSql
            End If
        End If
    Next x
End Sub

Private Sub Form_Load()
    Call PopulateDistrict
    Call SetGrid
End Sub

Private Sub Form_Resize()
    VSFlexGrid1.Height = Me.Height - 475
    VSFlexGrid1.Width = Me.Width - Picture1.Width - 325
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage True
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If SaleStaff_ Is Nothing Then Set SaleStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set SaleStaff_ = Nothing
        Set SaleStaff_ = New BWGCompanyStaff
    Else
        If Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set SaleStaff_ = Nothing
            Set SaleStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If SaleStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set SaleStaff_ = Nothing
    Set SaleStaff_ = selStaff
    If SaleStaff_ Is Nothing Then Exit Sub
    With SaleStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Private Sub PopulateDistrict()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colDistrict_ = tmpS.SaleDistrictSearch("isActive = 'Y'")
'    cboDistrict.Clear
'    cboDistrict.AddItem "== กรุณาระบุเขต Sale =="
    For x = 1 To colDistrict_.Count
        'cboDistrict.AddItem colDistrict_(x).DistrictName
        If x = 1 Then SDList = "#" & x & ";" & colDistrict_(x).ID & vbTab & colDistrict_(x).DistrictName & vbTab & colDistrict_(x).SaleStaffName
        If x > 1 Then SDList = SDList & "|#" & x & ";" & colDistrict_(x).ID & vbTab & colDistrict_(x).DistrictName & vbTab & colDistrict_(x).SaleStaffName
    Next
'    cboDistrict.ListIndex = 0
End Sub

Private Sub VSFlexGrid1_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
'    MsgBox VSFlexGrid1.ComboItem(VSFlexGrid1.ComboIndex)
End Sub
