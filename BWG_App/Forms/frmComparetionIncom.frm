VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmCompareIncome 
   Caption         =   "เปรียบเทียบยอดขาย"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   10695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCalcWaste 
      Caption         =   "Calc Waste"
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   4020
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   6375
      Left            =   2640
      TabIndex        =   15
      Top             =   0
      Width           =   7995
      _cx             =   14102
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
      BackColorAlternate=   12640495
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
      FixedRows       =   3
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
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   660
         TabIndex        =   16
         Top             =   4140
         Width           =   5775
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblPercent 
            Alignment       =   2  'Center
            Caption         =   "%"
            Height          =   195
            Left            =   2520
            TabIndex        =   18
            Top             =   600
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9960
      Left            =   0
      ScaleHeight     =   9930
      ScaleWidth      =   2625
      TabIndex        =   10
      Top             =   0
      Width           =   2655
      Begin VB.CheckBox ChkMinMaxWeight 
         Caption         =   "คำนวณขั้นต่ำ-สูง"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   4620
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox ChkTransport 
         Caption         =   "รวมค่าขนส่ง"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   4320
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CommandButton cmdVLookup 
         Caption         =   "VLookUP"
         Height          =   435
         Left            =   180
         TabIndex        =   19
         Top             =   7080
         Width           =   2295
      End
      Begin VB.CheckBox chkGroupTrips 
         Caption         =   "Goup by Trips"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   2940
         Width           =   2235
      End
      Begin VB.CheckBox chkGroupCust 
         Caption         =   "Goup by Customer"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   2640
         Width           =   2235
      End
      Begin VB.CheckBox chkGroupWaste 
         Caption         =   "Group by Waste"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   3540
         Width           =   2235
      End
      Begin VB.CheckBox chkGroupMenifest 
         Caption         =   "Goup by Menifest"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   3240
         Width           =   2235
      End
      Begin VB.ListBox lstCompany 
         Height          =   1635
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "ประมวลผล"
         Height          =   435
         Left            =   180
         TabIndex        =   8
         Top             =   6600
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   7680
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   6
         Top             =   5310
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   7
         Top             =   5910
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   3001
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2941
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เปรียบเทียบระหว่างบริษัท"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   13
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   5070
         Width           =   1830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   5670
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3001
         Y1              =   6360
         Y2              =   6360
      End
   End
End
Attribute VB_Name = "frmCompareIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer

Private Type SelectedCompany
    CompanyID As String
    ComapnyCode As String
    CompanyName As String
End Type
Dim colCompany() As SelectedCompany

Private Sub InitialSelection()
'On Error Resume Next
    Dim CalColumn As Integer, ComCount As Integer, YYCount As Integer, i As Integer
    CalColumn = 0
    ComCount = 0
    For i = 0 To lstCompany.ListCount - 1
        If lstCompany.Selected(i) = True Then
            ComCount = ComCount + 1
            ReDim colCompany(ComCount)
        End If
    Next i
    ComCount = 0
    For i = 0 To lstCompany.ListCount - 1
        If lstCompany.Selected(i) = True Then
            ComCount = ComCount + 1
            colCompany(ComCount - 1).CompanyID = CurrentUser.WorkCompanies(i + 1).ID
            colCompany(ComCount - 1).ComapnyCode = CurrentUser.WorkCompanies(i + 1).CompanyCode
            colCompany(ComCount - 1).CompanyName = CurrentUser.WorkCompanies(i + 1).CompanyNameTH
        End If
    Next i
    
    YYCount = dtTo.YearValue - dtFrom.YearValue + 1
    CalColumn = (ComCount * YYCount) + 3
    Call SetGrid(CalColumn, YYCount, dtFrom.YearValue, ComCount, chkCalcWaste.Value)
End Sub

Private Sub chkCalcWaste_Click()
    InitialSelection
End Sub

Private Sub chkCalcWaste_Validate(Cancel As Boolean)
    InitialSelection
End Sub

Private Sub cmdClose_Click()
    frmMain.Stb1.Panels(1).Text = App.Title
    Unload Me
End Sub

Private Sub cmdProcess_Click()
On Error GoTo ErrD
    Dim tSql As String, i As Integer
    Dim rsCust As New ADODB.Recordset, rsTrip As New ADODB.Recordset
    Dim rsCalc As New ADODB.Recordset, rsInsert As New ADODB.Recordset
    Dim CompanyID As String
    CompanyID = ""
    InitialSelection
    If MsgBox("การประมวลผลข้อมูลนี้ใช้เวลานานเป็นอย่างมาก" & vbCrLf & "เมื่อยืนยันการประมวลผลแล้วไม่ควรใช้คอมพิวเตอร์เครื่องนี้ทำงานอื่น" & vbCrLf & "เพราะจะทำให้ทรัพยากรไม่เพียงพอในการประมวลผลได้" & vbCrLf & vbCrLf & "ยืนยันการประมวลผล", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    Me.KeyPreview = False
    Frame1.Visible = True
    cmdProcess.Enabled = False
    cmdProcess.Caption = "กำลังประมวลผล"
    
    For i = 0 To UBound(colCompany) - 1
        If i = 0 Then CompanyID = "'" & colCompany(i).CompanyID & "'"
        If i > 0 Then CompanyID = CompanyID & ", '" & colCompany(i).CompanyID & "'"
    Next i
    
    
    DBConn.Execute "DELETE FROM tbCustomerIncome"
    
    '*** ค้นหาลูกค้าทั้งหมดที่มีการใช้บริการ
    
    '*** ค้นหาลูกค้าที่ขนของบริษัท AKP เท่านั้น
    DBConn.Execute "DROP TABLE tbCustomerAKPOnly"
    tSql = "SELECT DISTINCT CustomerID" & vbCrLf
    tSql = tSql & "INTO tbCustomerAKPOnly" & vbCrLf
    tSql = tSql & "FROM vw_CustomerInMenifest_ForCompare" & vbCrLf
    tSql = tSql & "WHERE WorkDate Between '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "'" & vbCrLf
    tSql = tSql & "AND CompanyID IN (" & CompanyID & ")"
'    tSql = tSql & "AND CompanyID IN ('776EDA5830F744A6B3551348D851BC22')" & vbCrLf
    tSql = tSql & "GROUP BY CustomerID" & vbCrLf
    DBConn.Execute tSql
    
    
    tSql = "SELECT DISTINCT CustomerCode, CustomerName, CustomerID, CompanyID, CompanyCode" & vbCrLf
    tSql = tSql & "FROM vw_CustomerInMenifest_ForCompare" & vbCrLf
    tSql = tSql & "WHERE WorkDate Between '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "'" & vbCrLf
'    tSql = tSql & "AND CompanyID IN (" & companyID & ")" & vbCrLf
    
    If Not curCustomer_ Is Nothing Then
        tSql = tSql & "AND CustomerID = '" & curCustomer_.ID & "'" & vbCrLf
    Else
        tSql = tSql & "AND CustomerID IN (SELECT CustomerID FROM tbCustomerAKPOnly)" & vbCrLf
    End If
'    tSql = tSql & "AND isCanceled <> 'Y' AND isClosed = 'Y'" & vbCrLf
    tSql = tSql & "GROUP BY CustomerName, CustomerCode, CustomerID, CompanyID, CompanyCode" & vbCrLf
    tSql = tSql & "ORDER BY CustomerName, CompanyCode"
    
    frmMain.Stb1.Panels(1).Text = "เริ่มการประมวลผลตามเงื่อนไขที่เลือก..."
    Set rsCust = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    frmMain.Stb1.Panels(1).Text = "ค้นพบข้อมูลลูกค้า " & rsCust.RecordCount & " รายการ"
    ProgressBar1.Max = rsCust.RecordCount
    For i = 1 To rsCust.RecordCount
        On Error Resume Next
        ProgressBar1.Value = i
        lblPercent.Caption = Format(i / rsCust.RecordCount * 100, "0.00") & "%"
        DoEvents
        frmMain.Stb1.Panels(1).Text = "กำลังค้นหาข้อมูลจำนวนเที่ยวเพื่อคำนวณรายได้"
        tSql = "SELECT DISTINCT TimeTableNo FROM Com_vwManifestWastAdd_Union" & vbCrLf
        tSql = tSql & "WHERE CustomerID = '" & rsCust!CustomerID & "' AND WorkDate >= '" & dtFrom.ValueYMD & "' AND WorkDate <= '" & dtTo.ValueYMD & "'" & vbCrLf
        tSql = tSql & "AND CompanyID =  '" & rsCust!CompanyID & "'" & vbCrLf
        tSql = tSql & "AND TimeTableNo IS NOT NULL"
        tSql = tSql & " AND isCanceled <> 'Y' AND isClosed = 'Y'" & vbCrLf
        tSql = tSql & "ORDER BY TimeTableNo"
        Set rsTrip = New ADODB.Recordset
        rsTrip.open tSql, DBConn, adOpenForwardOnly, adLockOptimistic
        If rsCust.RecordCount <= 0 Then GoTo ErrD
        frmMain.Stb1.Panels(1).Text = "ค้นหาข้อมูลจำนวนเที่ยวของ " & rsCust!CustomerName & " บริษัทที่รับกำจัด " & rsCust!CompanyCode & " จำนวน " & rsTrip.RecordCount & " เที่ยว"
        
        Dim curTrip As String
        Do While Not rsTrip.EOF
            DoEvents
'            Sleep 5000
            frmMain.Stb1.Panels(1).Text = "คำนวณรายได้ตามเที่ยวขน ใบคุมเลขที่ " & rsTrip!TimeTableNo
            
            tSql = "SELECT * FROM Com_vwManifestWastAdd_Union" & vbCrLf
            tSql = tSql & "WHERE TimeTableNo = '" & rsTrip!TimeTableNo & "'" & vbCrLf
            tSql = tSql & "AND CompanyID = '" & rsCust!CompanyID & "'" & vbCrLf
            tSql = tSql & "AND IsCanceled <> 'Y' AND isClosed = 'Y' AND isNull(Weight, 0) > 0" & vbCrLf
            tSql = tSql & "ORDER BY DocNo, WasteName"
            Set rsCalc = New ADODB.Recordset
            rsCalc.open tSql, DBConn, adOpenForwardOnly, adLockOptimistic
            frmMain.Stb1.Panels(1).Text = "ค้นพบรายการกากฯ ตามใบคุมเลขที่ " & rsTrip!TimeTableNo & " จำนวน " & rsCalc.RecordCount & " รายการ"
            DoEvents
            Dim Incomes As Double, j As Integer
            Dim SumWeightByTrip As Double, SumWeightForMinMax, MinWeight As Double, MaxWeight As Double
            Dim TransportFee As Double, MinWeightRate As Double, MaxWeightRate As Double, MaxWeightType As String
            Dim TreatmentUnit As String, NonMinWeight As Boolean, UnderRateIncome As Double, OverRateIncome As Double
            Incomes = 0
            SumWeightByTrip = 0
            TransportFee = 0
            TreatmentUnit = "บาท/ตัน"
                        
            MinWeight = 0
            MaxWeight = 0
            MinWeightRate = 0
            MaxWeightRate = 0
            MaxWeightType = ""
            SumWeightForMinMax = 0
            NonMinWeight = False
            UnderRateIncome = 0
            OverRateIncome = 0
            If rsCalc.RecordCount > 0 Then
                For j = 1 To rsCalc.RecordCount
                    If InStr(rsCalc!TreatmentUnitID, "เที่ยว") <> 0 Then
                        SumWeightByTrip = SumWeightByTrip + rsCalc!Weight
                    End If
                    
                    If rsCalc!IsPriceIncTransport = "N" Or rsCalc!IsPriceIncTransport <> "Y" Then
                        If TransportFee <= IIf(IsNull(rsCalc!TransportFee) = True, 0, rsCalc!TransportFee) Then
                            TransportFee = IIf(IsNull(rsCalc!TransportFee) = True, 0, rsCalc!TransportFee)
                        End If
                    End If
                    
                    SumWeightForMinMax = SumWeightForMinMax + rsCalc!Weight
                    
                    '*** ตรวจสอบ น้ำหนักขั้นต่ำ
                    If UCase(rsCalc!WasteType) = "NON" Or UCase(rsCalc!WasteType) = "NBL" Then
                        NonMinWeight = True
                        If MinWeight <= IIf(IsNull(rsCalc!MinWeightPerCar) = True, 0, rsCalc!MinWeightPerCar) Then
                            MinWeight = IIf(IsNull(rsCalc!MinWeightPerCar) = True, 0, rsCalc!MinWeightPerCar)
                        End If
                        If IsNull(rsCalc!NonMinRate) = True Then MinWeightRate = rsCalc!TreatmentRate
                        If rsCalc!NonMinRate = 0 Then
                            MinWeightRate = rsCalc!TreatmentRate
                        Else
                            If MinWeightRate <= IIf(IsNull(rsCalc!NonMinRate) = True, 0, rsCalc!NonMinRate) Then
                                MinWeightRate = IIf(IsNull(rsCalc!NonMinRate) = True, 0, rsCalc!NonMinRate)
                            End If
                        End If
                    Else
                        If NonMinWeight = False Then
                            If MinWeight <= IIf(IsNull(rsCalc!MinWeightPerCar) = True, 0, rsCalc!MinWeightPerCar) Then
                                MinWeight = IIf(IsNull(rsCalc!MinWeightPerCar) = True, 0, rsCalc!MinWeightPerCar)
                            End If
                            If MinWeightRate <= IIf(IsNull(rsCalc!TreatmentRate) = True, 0, rsCalc!TreatmentRate) Then
                                MinWeightRate = IIf(IsNull(rsCalc!TreatmentRate) = True, 0, rsCalc!TreatmentRate)
                            End If
                        End If
                    End If
                    
                    '*** ตรวจสอบ น้ำหนักขั้นสูง
                    If IsNull(rsCalc!MaxWeight) = True Then MaxWeight = 0
                    If rsCalc!MaxWeight > 0 Then
                        If InStr(rsCalc!MaxRateType, "เที่ยว") <> 0 Then
                            MaxWeightType = rsCalc!MaxRateType
                            If MaxWeight <= IIf(IsNull(rsCalc!MaxWeight) = True, 0, rsCalc!MaxWeight) Then
                                MaxWeight = IIf(IsNull(rsCalc!MaxWeight) = True, 0, rsCalc!MaxWeight)
                            End If
                            If MaxWeightRate <= IIf(IsNull(rsCalc!MaxRate) = True, 0, rsCalc!MaxRate) Then
                                MaxWeightRate = IIf(IsNull(rsCalc!MaxRate) = True, 0, rsCalc!MaxRate)
                            End If
                        Else
                            If MaxWeightType <> "บาท/เที่ยว" Then
                                If MaxWeight <= IIf(IsNull(rsCalc!MaxWeight) = True, 0, rsCalc!MaxWeight) Then
                                    MaxWeight = IIf(IsNull(rsCalc!MaxWeight) = True, 0, rsCalc!MaxWeight)
                                End If
                                If MaxWeightRate <= IIf(IsNull(rsCalc!MaxRate) = True, 0, rsCalc!MaxRate) Then
                                    MaxWeightRate = IIf(IsNull(rsCalc!MaxRate) = True, 0, rsCalc!MaxRate)
                                End If
                            End If
                        End If
                    End If
                                        
                    rsCalc.MoveNext
                Next j
                
                If MinWeight > 0 Then
                    If SumWeightForMinMax < MinWeight Then
                        UnderRateIncome = (MinWeight - SumWeightForMinMax) * MinWeightRate
                    Else
                        UnderRateIncome = 0
                    End If
                End If
                If MaxWeight > 0 Then
                    If SumWeightForMinMax > MaxWeight Then
                        If MaxWeightType = "บาท/เที่ยว" Then
                            OverRateIncome = MaxWeightRate
                        Else
                            OverRateIncome = (SumWeightForMinMax - MaxWeight) * MaxWeightRate
                        End If
                    End If
                End If
                
                rsCalc.MoveFirst
                Do While Not rsCalc.EOF
                    DoEvents
                    '*** คำนวณยอดขาย
                    frmMain.Stb1.Panels(1).Text = "คำนวณยอดขาย กากของเสีย " & rsCalc!WasteName
                    If InStr(rsCalc!TreatmentUnitID, "เที่ยว") <> 0 Then
                        frmMain.Stb1.Panels(1).Text = frmMain.Stb1.Panels(1).Text & " (เที่ยว) " & rsCalc!TreatmentRate * (rsCalc!Weight * 100 / SumWeightByTrip) / 100
                        Incomes = rsCalc!TreatmentRate * (rsCalc!Weight * 100 / SumWeightByTrip) / 100 ' ไม่รวมค่าขนส่ง
                    Else
                        frmMain.Stb1.Panels(1).Text = frmMain.Stb1.Panels(1).Text & " " & rsCalc!TreatmentRate * rsCalc!Weight ' ไม่รวมค่าขนส่ง
                        Incomes = rsCalc!TreatmentRate * rsCalc!Weight
                    End If
                    
                    '*** คำนวนยอดขายรวมค่าขนส่งเฉลี่ยตามเวสต์
                    If ChkTransport.Value = 1 Then
                        If TransportFee > 0 Then Incomes = Incomes + (TransportFee / rsCalc.RecordCount) ' ยอดขาย + ค่าขนส่งเฉลี่ยตามเวสต์
                    End If
                    
                    '*** คำนวนยอดขายรวมน้ำหนักขั้นต่ำ-สูง
                    If ChkMinMaxWeight.Value = 1 Then
                        If MinWeight > 0 Then
                            If UnderRateIncome > 0 Then Incomes = Incomes + (UnderRateIncome / rsCalc.RecordCount)
                        End If
                        If MaxWeight > 0 Then
                            If OverRateIncome > 0 Then Incomes = Incomes + (OverRateIncome / rsCalc.RecordCount)
                        End If
                    End If
                    '**********************************
                    
                    
                    
                    DoEvents
                    frmMain.Stb1.Panels(1).Text = "บันทึกรายได้หลังจากการคำนวณลงในฐานข้อมูล tbCustomerIncome"
                    tSql = "SELECT CustomerID, CustomerCode, CustomerName, WasteName, WorkYear, CompanyCode, Incomes" & vbCrLf
                    tSql = tSql & "FROM tbCustomerIncome" & vbCrLf
                    tSql = tSql & "WHERE CustomerID = '" & rsCalc!CustomerID & "' AND WasteName = '" & rsCalc!WasteName & "' AND WorkYear = " & Left(rsCalc!WorkDate, 4) & " AND CompanyCode = '" & rsCust!CompanyCode & "'"
                    Set rsInsert = New ADODB.Recordset
                    rsInsert.open tSql, DBConn, adOpenDynamic, adLockOptimistic
                    If rsInsert.EOF Then
                        rsInsert.AddNew
                        rsInsert!CustomerID = rsCalc!CustomerID
                        rsInsert!CustomerCode = rsCust!CustomerCode
                        rsInsert!CustomerName = rsCust!CustomerName
                        rsInsert!WasteName = rsCalc!WasteName
                        rsInsert!WorkYear = Left(rsCalc!WorkDate, 4)
                        rsInsert!CompanyCode = rsCust!CompanyCode
                        rsInsert!Incomes = Incomes
                    Else
                        rsInsert!Incomes = rsInsert!Incomes + Incomes
                    End If
                    rsInsert.Update
                    frmMain.Stb1.Panels(1).Text = "บันทึกสำเร็จ (" & rsCust!CustomerName & " รายการของเสีย " & rsCalc!WasteName & " ปี " & Left(rsCalc!WorkDate, 4) & " บริษัทที่รับ " & rsCust!CompanyCode

                    '*** เขียนข้อมูลลงใน Grid
                    WriteGrid rsCust!CustomerCode, rsCust!CustomerName, Replace(rsCalc!WasteName, vbTab, ""), Left(rsCalc!WorkDate, 4), rsCust!CompanyCode, Incomes

'                    frmMain.Stb1.Panels(2).Text = "Rows = " & VSFlexGrid1.Rows - 1
'                    VSFlexGrid1.TopRow = VSFlexGrid1.BottomRow
'                    VSFlexGrid1.Refresh
                    DoEvents
                    Set rsInsert = Nothing
                    
                    rsCalc.MoveNext
                Loop
                Set rsCalc = Nothing
            
            End If
            
            rsTrip.MoveNext
        Loop
        Set rsTrip = Nothing
        
        rsCust.MoveNext
    Next i
    Set rsCust = Nothing
    
    Frame1.Visible = False
    frmMain.Stb1.Panels(1).Text = "การประมวลผลเสร็จสิ้น"
'    txtLog.Text = txtLog.Text & vbCrLf & frmMain.Stb1.Panels(1).Text
    Me.KeyPreview = True
    cmdProcess.Enabled = True
    cmdProcess.Caption = "ประมวลผล"
    
    VSFlexGrid1.SaveGrid App.Path & "\" & "VLookUP.xls", flexFileTabText
    
    'First lets get started on opening the excel sheet itself
    Dim xlsApp, xlsBook
    Set xlsApp = CreateObject("Excel.Application") 'creates the excel
    xlsApp.Visible = True 'show the excel form
    'xlsApp.Workbooks.Add 'add an excel workbook
    Set xlsBook = xlsApp.Workbooks.open(App.Path & "\" & "VLookUP.xls")

    Set rsCust = Nothing
    Exit Sub
ErrD:
    frmMain.Stb1.Panels(1).Text = err.Description
    Resume Next
End Sub

Private Sub cmdVLookup_Click()
    Dim tSql As String
    Dim RS As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    tSql = "SELECT * FROM tbCustomerIncome ORDER BY CustomerName, CustomerCode, WasteName, WorkYear, CompanyCode"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            WriteGrid RS!CustomerCode, RS!CustomerName, Replace(RS!WasteName, vbTab, ""), RS!WorkYear, RS!CompanyCode, RS!Incomes
            RS.MoveNext
        Loop
        Set RS = Nothing
        VSFlexGrid1.SaveGrid App.Path & "\" & "VLookUP.xls", flexFileTabText
        
        'First lets get started on opening the excel sheet itself
        Dim xlsApp, xlsBook
        Set xlsApp = CreateObject("Excel.Application") 'creates the excel
        xlsApp.Visible = True 'show the excel form
        Set xlsBook = xlsApp.Workbooks.open(App.Path & "\" & "VLookUP.xls")
        
    Else
        MsgBox "ไม่พบข้อมูล", vbInformation
    End If
    
End Sub

Private Sub WriteGrid(CustCode As String, custName As String, WasteName As String, WorkYear As Integer, CompanyCode As String, Incomes As Double)
    '*** เขียนข้อมูลลงใน Grid
    Dim RowCust As Long, RowWaste As Long, yCol As Integer, cCol As Integer
    Dim DecimalValue As Double, RowKey As String
    Dim iCol As Long, CalcWaste As Boolean
    DecimalValue = Incomes - Int(Incomes)
    If DecimalValue > 0 Then
'        MsgBox DecimalValue
    End If
    CalcWaste = chkCalcWaste.Value
    If CalcWaste Then
        RowKey = CustCode & "_" & Trim(WasteName)
        iCol = 3
    Else
        RowKey = CustCode
        iCol = 2
    End If
        
    frmMain.Stb1.Panels(2).Text = "Rows = " & VSFlexGrid1.Rows - 1
    DoEvents
    
    RowCust = VSFlexGrid1.FindRow(RowKey, , 0, False)
    If RowCust = -1 Then
        If CalcWaste Then VSFlexGrid1.AddItem RowKey & vbTab & custName & vbTab & Trim(WasteName)
        If CalcWaste = False Then VSFlexGrid1.AddItem RowKey & vbTab & custName
        RowCust = VSFlexGrid1.FindRow(RowKey, , 0, False)
        For yCol = iCol To VSFlexGrid1.Cols - 1
            If WorkYear = VSFlexGrid1.TextMatrix(1, yCol) Then
                For cCol = yCol To VSFlexGrid1.Cols - 1
                    If (WorkYear = VSFlexGrid1.TextMatrix(1, yCol)) And (CompanyCode = VSFlexGrid1.TextMatrix(2, cCol)) Then
                        VSFlexGrid1.TextMatrix(RowCust, cCol) = Format(Incomes, "#,##0.00")
                        Exit For
                    End If
                Next cCol
                Exit For
            End If
        Next yCol
        VSFlexGrid1.TopRow = RowCust
        VSFlexGrid1.Refresh
        'DoEvents
    Else
        RowWaste = VSFlexGrid1.FindRow(RowKey, , 0, False)
        If RowWaste = -1 Then
            If CalcWaste Then VSFlexGrid1.AddItem RowKey & vbTab & custName & vbTab & Trim(WasteName)
            If CalcWaste = False Then VSFlexGrid1.AddItem RowKey & vbTab & custName
            RowWaste = VSFlexGrid1.FindRow(RowKey, , 0, False)
        End If
        
        For yCol = iCol To VSFlexGrid1.Cols - 1
            If WorkYear = VSFlexGrid1.TextMatrix(1, yCol) Then
                For cCol = yCol To VSFlexGrid1.Cols - 1
                    If (WorkYear = VSFlexGrid1.TextMatrix(1, yCol)) And (CompanyCode = VSFlexGrid1.TextMatrix(2, cCol)) Then
                        VSFlexGrid1.TextMatrix(RowWaste, cCol) = Format(CDbl(Val(Format(VSFlexGrid1.TextMatrix(RowWaste, cCol), "0.00"))) + Incomes, "#,##0.00")
                        Exit For
                    End If
                Next cCol
                Exit For
            End If
        Next yCol
        VSFlexGrid1.TopRow = RowWaste
        VSFlexGrid1.Refresh
        'DoEvents
    End If
End Sub

Private Sub dtFrom_Validate(Cancel As Boolean)
    Call InitialSelection
End Sub

Private Sub dtTo_Validate(Cancel As Boolean)
    Call InitialSelection
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Frame1.Visible = False
    PopulateCompany lstCompany, 0
    If lstCompany.ListCount > 2 Then lstCompany.RemoveItem 0
    Call InitialSelection
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Container.Name = dtFrom.Container.Name Or Me.ActiveControl.Container.Name = dtTo.Container.Name Then
            Call InitialSelection
        End If
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With VSFlexGrid1
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 100
        End If
        .Height = (Me.Height - .Top) - 500
    End With
    Frame1.Left = (Me.Width / 2) - Frame1.Width
'    Call InitialSelection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub SetGrid(Optional ByVal ColCount As Integer = 9, Optional ByVal YYCount As Integer = 3, Optional FirstYear As Integer = 2550, Optional ByVal ComCount As Integer = 2, Optional ByVal CalcWaste As Boolean = True)
    On Error Resume Next
    Dim i As Integer, j As Integer
    Dim iCol As Integer
    With VSFlexGrid1
        .Clear
        .Rows = 3
        .Cols = ColCount
        .FixedCols = 1
        .FixedRows = 3
        .MergeCells = flexMergeFixedOnly
        .MergeRow(-1) = True
        .MergeCol(-1) = True
        .CellAlignment = flexAlignCenterCenter
        
        For i = .FixedCols To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        
        .Cell(flexcpText, 0, 0, 2, 0) = "CustCode"
        .Cell(flexcpText, 0, 1, 2, 1) = "Customer"
        
        If CalcWaste Then
            iCol = 3
            .Cols = ColCount
            .Cell(flexcpText, 0, 2, 2, 2) = "Waste"
            .ColWidth(0) = 0
            .ColWidth(1) = (.Width - (.ColWidth(iCol) * (.Cols - iCol))) / 2
            .ColWidth(2) = .ColWidth(1)
        Else
            iCol = 2
            .Cols = ColCount - 1
            .ColWidth(0) = 0
            .ColWidth(iCol) = 960
            .ColWidth(1) = (.Width - (.ColWidth(iCol) * (.Cols - iCol)))
            '.ColWidth(2) = .ColWidth(1)
        End If
        .Select 0, iCol, 0, .Cols - 1
        .Cell(flexcpText, 0, iCol, 1, .Cols - 1) = "Incomes"
        
        For i = 0 To YYCount - 1
            .Cell(flexcpText, 1, iCol + (i * ComCount), 1, iCol + (ComCount - 1) + (i * ComCount)) = FirstYear + i
            For j = 0 To ComCount - 1
                .TextMatrix(2, iCol + (i * ComCount) + j) = colCompany(j).ComapnyCode
            Next j
        Next i
                    
        For i = iCol To .Cols - 1
            .ColDataType(i) = flexDTDouble
        Next i
    End With
End Sub

Private Sub lstCompany_ItemCheck(Item As Integer)
    Call InitialSelection
End Sub

Private Sub lstCompany_Validate(Cancel As Boolean)
    Call InitialSelection
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If txtCustomer = "*" Then Set curCustomer_ = Nothing: Exit Sub
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If Trim(curCustomer_.ID) = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub

