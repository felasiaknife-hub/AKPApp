VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTransportChargRate 
   Caption         =   "กำหนดอัตราค่าขนส่ง"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   12600
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   9315
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   9015
      _cx             =   15901
      _cy             =   16431
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   2655
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export To XLS"
         Height          =   435
         Left            =   420
         TabIndex        =   3
         Top             =   3240
         Width           =   1755
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   420
         TabIndex        =   4
         Top             =   3900
         Width           =   1755
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   420
         TabIndex        =   2
         Top             =   2580
         Width           =   1755
      End
      Begin VB.ComboBox cboTruckType 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1500
         Width           =   2115
      End
      Begin VB.TextBox txtFromProvince 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   300
         TabIndex        =   0
         Text            =   "*"
         Top             =   780
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "เลือก Period"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "จังหวัด"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmTransportChargRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim coltruckType_ As Collection
Dim curFromProvince_ As IMIProvince
Dim curToProvince_ As IMIProvince
Dim provLocation_ As String
Dim TransRateType As String

Public Sub TransType(ByVal Value As String)
    TransRateType = Value
End Sub

Private Sub SetGrid()
Dim AlignRight As Integer
    With VSFlexGrid1
        .Rows = 50
        .Cols = 16
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 2300
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 1300
        .ColWidth(7) = 1300
        .ColWidth(8) = 1300
        .ColWidth(9) = 1300
        .ColWidth(10) = 1300
        .ColWidth(11) = 1300
        .ColWidth(12) = 0
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "ต้นทาง"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = ""
        .TextMatrix(0, 5) = ""
        .TextMatrix(0, 6) = ""
        .TextMatrix(0, 7) = ""
        .TextMatrix(0, 8) = ""
        .TextMatrix(0, 9) = ""
        .TextMatrix(0, 10) = ""
        .TextMatrix(0, 11) = ""
        .TextMatrix(0, 12) = ""
        
        .FixedCols = 3
        .FixedRows = 2
        AlignRight = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = AlignRight
        .ColAlignment(4) = AlignRight
        .ColAlignment(5) = AlignRight
        .ColAlignment(6) = AlignRight
        .ColAlignment(7) = AlignRight
        .ColAlignment(8) = AlignRight
        .ColAlignment(9) = AlignRight
        .ColAlignment(10) = AlignRight
        .ColAlignment(11) = AlignRight
        .ColAlignment(12) = AlignRight
        .AllowUserResizing = flexResizeColumns
    End With
End Sub

Private Sub GetTruckHeader()
Dim RS As ADODB.Recordset
Dim tSql As String, i As Integer
    tSql = "SELECT * FROM tbTruckType WHERE isActive = 'Y'" & vbCrLf
    tSql = tSql & "ORDER BY TruckTypeDesc"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If RS.RecordCount = 0 Then Exit Sub
    VSFlexGrid1.Cols = RS.RecordCount + 6
    i = 3
    Do While Not RS.EOF
        VSFlexGrid1.TextMatrix(0, 0) = ""
        VSFlexGrid1.TextMatrix(0, 1) = ""
        If RS!TruckTypeID = "DA383A0BDFFF4FD09BA97EB016C919AE" Then
            VSFlexGrid1.TextMatrix(0, i) = RS!TruckTypeID
            VSFlexGrid1.TextMatrix(1, i) = RS!TruckTypeDesc
            If i >= 3 Then VSFlexGrid1.ColWidth(i) = 1200
            i = i + 1
            VSFlexGrid1.TextMatrix(0, i) = RS!TruckTypeID
            VSFlexGrid1.TextMatrix(1, i) = RS!TruckTypeDesc & " - พ่วง"
        ElseIf RS!TruckTypeID = "67D1CCC9C8FF4782B89E3F56100A1592" Then
            VSFlexGrid1.TextMatrix(0, i) = RS!TruckTypeID
            VSFlexGrid1.TextMatrix(1, i) = RS!TruckTypeDesc
            If i >= 3 Then VSFlexGrid1.ColWidth(i) = 1200
            i = i + 1
            VSFlexGrid1.TextMatrix(0, i) = RS!TruckTypeID
            VSFlexGrid1.TextMatrix(1, i) = RS!TruckTypeDesc & " - พ่วง"
        Else
            VSFlexGrid1.TextMatrix(0, i) = RS!TruckTypeID
            VSFlexGrid1.TextMatrix(1, i) = RS!TruckTypeDesc
        End If
        If i >= 3 Then VSFlexGrid1.ColWidth(i) = 1200
        RS.MoveNext
        i = i + 1
    Loop
    RS.Close
    Set RS = Nothing
    VSFlexGrid1.RowHeight(0) = 0
    VSFlexGrid1.ColWidth(VSFlexGrid1.Cols - 1) = 0
End Sub

Private Sub GetTruckChargRate(ByVal TrnView As String)
On Error Resume Next
Dim RS As ADODB.Recordset
Dim tSql As String, i As Integer, j As Integer
Dim TruckID_Header As String
Dim Period() As String
    frmMain.Stb1.Panels(1).Text = ""
    tSql = "SELECT * FROM tbTruckType WHERE isActive = 'Y'" & vbCrLf
    tSql = tSql & "ORDER BY TruckTypeDesc"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    TruckID_Header = ""
    If RS.RecordCount = 0 Then Exit Sub
    If Trim(txtFromProvince.Text) = "" Or txtFromProvince.Text = "*" Then Set curFromProvince_ = Nothing
    Do While Not RS.EOF
        If TruckID_Header = "" Then
            If RS!TruckTypeID = "DA383A0BDFFF4FD09BA97EB016C919AE" Then
                TruckID_Header = "MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN RatePerTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "]'" & vbCrLf
                TruckID_Header = TruckID_Header & ", SUM(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN TrailerRateTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "- พ่วง]'" & vbCrLf
            ElseIf RS!TruckTypeID = "67D1CCC9C8FF4782B89E3F56100A1592" Then
                TruckID_Header = "MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN RatePerTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "]'" & vbCrLf
                TruckID_Header = TruckID_Header & ", SUM(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN TrailerRateTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "- พ่วง]'" & vbCrLf
            Else
                TruckID_Header = "MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN RatePerTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "]'" & vbCrLf
            End If
        Else
            If RS!TruckTypeID = "DA383A0BDFFF4FD09BA97EB016C919AE" Then
                TruckID_Header = TruckID_Header & ", MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN RatePerTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "]'" & vbCrLf
                TruckID_Header = TruckID_Header & ", MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN TrailerRateTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "- พ่วง]'" & vbCrLf
            ElseIf RS!TruckTypeID = "67D1CCC9C8FF4782B89E3F56100A1592" Then
                TruckID_Header = TruckID_Header & ", MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN RatePerTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "]'" & vbCrLf
                TruckID_Header = TruckID_Header & ", MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN TrailerRateTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "- พ่วง]'" & vbCrLf
            Else
                TruckID_Header = TruckID_Header & ", MAX(CASE WHEN CarTypeID = '" & Trim(RS!TruckTypeID) & "' THEN RatePerTrip ELSE 0 END) AS '[" & RS!TruckTypeDesc & "]'" & vbCrLf
            End If
        End If
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    
    If Trim(TruckID_Header) = "" Then Exit Sub
    tSql = "SELECT FromProvinceID, FromAumphurID, FromProvinceName + ' - ' + FromAumphurName + ' (' + ToProvinceName + ')' AS Desctination," & vbCrLf
    tSql = tSql & TruckID_Header
'    tSql = tSql & ", Max(ActiveRate) AS ActiveRate" & vbCrLf
    tSql = tSql & "From " & TrnView & vbCrLf 'vw_TransportCustChargeRate" & vbCrLf
    If Trim(txtFromProvince) <> "" Or Trim(txtFromProvince) <> "*" Then
        If Not curFromProvince_ Is Nothing Then
            tSql = tSql & "WHERE FromProvinceID = '" & curFromProvince_.ID & "'" & vbCrLf
        End If
    End If
    
    If cboTruckType.Text <> "" Then
        Period = Split(cboTruckType.Text, "-")
        If InStr(tSql, "WHERE") > 0 Then
            tSql = tSql & "AND StartDate = '" & Trim(Period(0)) & "' AND EndDate = '" & Trim(Period(1)) & "'" & vbCrLf
        Else
            tSql = tSql & "WHERE StartDate = '" & Trim(Period(0)) & "' AND EndDate = '" & Trim(Period(1)) & "'" & vbCrLf
        End If
    End If
    
    tSql = tSql & "GROUP BY ToProvinceName, FromProvinceName, FromAumphurName, FromProvinceID, FromAumphurID"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    TruckID_Header = ""
    If RS.RecordCount = 0 Then Exit Sub
    VSFlexGrid1.Rows = RS.RecordCount + 3
    i = 2
    Do While Not RS.EOF
        With VSFlexGrid1
            For j = 0 To RS.Fields.Count - 1
                .TextMatrix(i, j) = RS.Fields(j).Value
            Next j
        End With
        RS.MoveNext
        i = i + 1
    Loop
    frmMain.Stb1.Panels(1).Text = "ค้นหาข้อมูลอัตราค่าขนส่ง " & txtFromProvince & " พบ " & RS.RecordCount & " รายการ"
    RS.Close
    Set RS = Nothing
End Sub

Private Sub GetPeriod(ByVal TrnView As String)
Dim RS As ADODB.Recordset
Dim x%, tSql As String
        
    tSql = "SELECT DISTINCT StartDate + '-' + EndDate AS Period FROM " & TrnView & " ORDER BY Period"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    cboTruckType.Clear
    cboTruckType.AddItem "== กรุณาระบุ ช่วงวันที่ =="
    Do While Not RS.EOF
        cboTruckType.AddItem RS!Period
        RS.MoveNext
    Loop
    Set RS = Nothing
    cboTruckType.ListIndex = 0
End Sub

Private Sub cmdExport_Click()
Dim AppExcel As Object
Dim ExcelFile As String
    Set AppExcel = CreateObject("Excel.Application")
    ExcelFile = App.Path & "\" & TransRateType & IIf(cboTruckType.ListIndex > 0, "_" & Replace(cboTruckType.Text, "/", "-"), "") & "_" & Replace(Format(Date, "yyyy/MM/dd") & "_" & Format(Time, "HH-MM"), "/", "-") & ".xls"
    VSFlexGrid1.FixedCols = 2
    VSFlexGrid1.FixedRows = 1
    VSFlexGrid1.SaveGrid ExcelFile, flexFileTabText
    AppExcel.Workbooks.open ExcelFile
    
    AppExcel.Visible = True
    Set AppExcel = Nothing
    
    VSFlexGrid1.FixedCols = 3
    VSFlexGrid1.FixedRows = 2
    
End Sub

Private Sub cmdSearch_Click()
    VSFlexGrid1.Clear
    SetGrid
    GetTruckHeader
    GetTruckChargRate TransRateType
End Sub

Private Sub Command1_Click()
    frmMain.Stb1.Panels(1).Text = ""
    Unload Me
End Sub

Private Sub VSFlexGrid1_BeforeUpdateValue(NewValue As String, Row As Integer, Col As Integer, Cancel As Boolean)
    If IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ค่าเป็นตัวเลขเท่านั้น", vbCritical
        Cancel = True
    End If
End Sub

Private Sub VSFlexGrid1_UpdateValue(Row As Integer, Col As Integer)
'On Error GoTo ErrD
Dim ProvinceID As String, AumphurID As String, TruckTypeID As String
Dim tSql As String, tmpSql As String, RS As New ADODB.Recordset
Dim ActiveRate As Integer, TransRate As Double, isTrailer As Boolean
Dim TruckTypeDesc As String
    ProvinceID = VSFlexGrid1.TextMatrix(Row, 0)
    AumphurID = VSFlexGrid1.TextMatrix(Row, 1)
    TruckTypeID = VSFlexGrid1.TextMatrix(0, Col)
    TruckTypeDesc = VSFlexGrid1.TextMatrix(1, Col)
    ActiveRate = VSFlexGrid1.TextMatrix(Row, VSFlexGrid1.Cols - 1)
    TransRate = VSFlexGrid1.TextMatrix(Row, Col)
    isTrailer = False
    If TruckTypeID = "DA383A0BDFFF4FD09BA97EB016C919AE" Then
        If InStr(1, TruckTypeDesc, "พ่วง") <> 0 Then
            isTrailer = True
        End If
    ElseIf TruckTypeID = "67D1CCC9C8FF4782B89E3F56100A1592" Then
        If InStr(1, VSFlexGrid1.TextMatrix(1, Col), "พ่วง") > 0 Then
            isTrailer = True
        End If
    End If
    If ActiveRate = 0 Then
        If isTrailer = False Then
            tSql = "UPDATE tbTransVendorChargeRate SET RatePerTrip = " & TransRate & ", CurRatePerTrip = " & TransRate
        Else
            tSql = "UPDATE tbTransVendorChargeRate SET TrailerRateTrip = " & TransRate & ", CurTrailerRateTrip = " & TransRate
        End If
    ElseIf ActiveRate = 1 Then
        If isTrailer = False Then
            tSql = "UPDATE tbTransVendorChargeRate SET RatePerTrip1 = " & TransRate & ", CurRatePerTrip = " & TransRate
        Else
            tSql = "UPDATE tbTransVendorChargeRate SET TrailerRateTrip1 = " & TransRate & ", CurTrailerRateTrip = " & TransRate
        End If
    ElseIf ActiveRate = 2 Then
        If isTrailer = False Then
            tSql = "UPDATE tbTransVendorChargeRate SET RatePerTrip2 = " & TransRate & ", CurRatePerTrip = " & TransRate
        Else
            tSql = "UPDATE tbTransVendorChargeRate SET TrailerRateTrip2 = " & TransRate & ", CurTrailerRateTrip = " & TransRate
        End If
    ElseIf ActiveRate = 3 Then
        If isTrailer = False Then
            tSql = "UPDATE tbTransVendorChargeRate SET RatePerTrip3 = " & TransRate & ", CurRatePerTrip = " & TransRate
        Else
            tSql = "UPDATE tbTransVendorChargeRate SET TrailerRateTrip3 = " & TransRate & ", CurTrailerRateTrip = " & TransRate
        End If
    End If
    tSql = tSql & vbCrLf & "WHERE FromAumphurID = '" & AumphurID & "' AND CarTypeID = '" & TruckTypeID & "'"
    
    tmpSql = "SELECT * FROM tbTransVendorChargeRate "
    tmpSql = tmpSql & vbCrLf & "WHERE FromAumphurID = '" & AumphurID & "' AND CarTypeID = '" & TruckTypeID & "'"
    Set RS = New ADODB.Recordset
    RS.open tmpSql, DBConn, adOpenForwardOnly, adLockReadOnly
    If RS.RecordCount > 0 Then
        DBConnExc tSql
    Else
        tSql = "INSERT INTO tbTransVendorChargeRate( ChargeRateID, CarTypeID, FromAumphurID, ToAumphurID, RatePerKM, RatePerTrip, EstDistance, LastUpdateDate, ChargeSubID, CustRatePerKM, RatePerTrip," & vbCrLf
        tSql = tSql & "                      RatePerVol, CustRatePerVol, TrailerRateKM, TrailerRateTrip, TrailerRateVol, TrailerCustRateKM, TrailerRateTrip, TrailerCustRateVol," & vbCrLf
        tSql = tSql & "                      RatePerTrip1, RatePerTrip2, RatePerTrip3, TrailerRateTrip1, TrailerRateTrip2, TrailerRateTrip3, CurRatePerTrip," & vbCrLf
        tSql = tSql & "                      CurTrailerRateTrip, ActiveRate)" & vbCrLf
        tSql = tSql & "SELECT  Replace(NEWID(), '-', ''), CarTypeID, '" & AumphurID & "' AS FromAumphurID, ToAumphurID, 0 AS RatePerKM, 0 AS RatePerTrip, 0 AS EstDistance, '" & Format(Now, "yyyy/mm/dd") & "' AS LastUpdateDate, '001' AS ChargeSubID, 0 AS CustRatePerKM, 0 AS RatePerTrip," & vbCrLf
        tSql = tSql & "                      0 AS RatePerVol, 0 AS CustRatePerVol, 0 AS TrailerRateKM, 0 AS TrailerRateTrip, 0 AS TrailerRateVol, 0 AS TrailerCustRateKM, 0 AS TrailerRateTrip, 0 AS TrailerCustRateVol," & vbCrLf
        tSql = tSql & "                      0 AS RatePerTrip1, 0 AS RatePerTrip2, 0 AS RatePerTrip3, 0 AS TrailerRateTrip1, 0 AS TrailerRateTrip2, 0 AS TrailerRateTrip3, 0 AS CurRatePerTrip," & vbCrLf
        tSql = tSql & "                      0 AS CurTrailerRateTrip , ActiveRate" & vbCrLf
        tSql = tSql & "From tbTransVendorChargeRate" & vbCrLf
        tSql = tSql & "WHERE FromAumphurID = 'AP20061229120117' AND CarTypeID = '" & TruckTypeID & "'"
        DBConnExc tSql
    End If
    
    Exit Sub
ErrD:
    MsgBox err.Description
    Resume Next
End Sub

Private Sub Form_Load()
    VSFlexGrid1.Clear
     Call GetPeriod(TransRateType)
    Call SetGrid
    Call GetTruckHeader
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Frame1.Height = Me.Height
    VSFlexGrid1.Width = Me.Width - Frame1.Width - 200
    VSFlexGrid1.Height = Me.Height - frmMain.Stb1.Height - 100
End Sub

Public Sub SetDataProvince(selProvince As IMIProvince, Optional isAssigned As Boolean = True)
    If provLocation_ = "F" Then
        Set curFromProvince_ = selProvince
        With curFromProvince_
            txtFromProvince.Text = .ProvinceName
            If isAssigned Then
'                myMenifest_.FromProvinceID = .ID
'                myMenifest_.FromProvinceName = .ProvinceName
            End If
        End With
'    ElseIf provLocation_ = "T" Then
'        Set curToProvince_ = selProvince
'        With curToProvince_
'            txtToProvince.Text = .ProvinceName
'            If isAssigned Then
'                myMenifest_.ToProvinceID = .ID
'                myMenifest_.ToProvinceName = .ProvinceName
'            End If
'        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub txtFromProvince_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtFromProvince_Validate False
End Sub

Private Sub txtFromProvince_Validate(Cancel As Boolean)
    If Trim(txtFromProvince.Text) = "*" Then Exit Sub
    If Trim(txtFromProvince.Text) = "" Then Exit Sub
    If curFromProvince_ Is Nothing Then Set curFromProvince_ = New IMIProvince
    If Trim(txtFromProvince.Text) = "" And Trim(curFromProvince_.ProvinceName) <> Trim(txtFromProvince.Text) Then
        Set curFromProvince_ = Nothing
        Set curFromProvince_ = New IMIProvince
'        With myMenifest_
'            .FromProvinceID = ""
'            .FromProvinceName = ""
'        End With
    Else
        If Trim(curFromProvince_.ProvinceName) <> Trim(txtFromProvince.Text) Then
            Set curFromProvince_ = Nothing
            Set curFromProvince_ = New IMIProvince
'            With myMenifest_
'                .FromProvinceID = ""
'                .FromProvinceName = ""
'            End With
            provLocation_ = "F"
            Load frmProvinceSearch
            With frmProvinceSearch
                .getCriteria (txtFromProvince.Text)
                .SetCallerForm Me
                .Show 1
            End With
'            If myMenifest_.FromProvinceID = "" Then txtFromProvince.Text = ""
        End If
    End If
End Sub
