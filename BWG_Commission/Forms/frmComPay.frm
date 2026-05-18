VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComPay 
   Caption         =   "Commission Payment"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   2640
      TabIndex        =   17
      Top             =   -120
      Width           =   9675
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   660
         TabIndex        =   25
         Top             =   780
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "พิมพ์ค่าคอมฯ"
         Height          =   435
         Left            =   660
         TabIndex        =   12
         Top             =   9000
         Visible         =   0   'False
         Width           =   2115
      End
      Begin Commission.ctlGrid ctlGrid1 
         Height          =   2595
         Left            =   660
         TabIndex        =   23
         Top             =   1200
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4577
      End
      Begin Commission.ctlGrid ctlGrid2 
         Height          =   2235
         Left            =   660
         TabIndex        =   24
         Top             =   3960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3942
      End
      Begin Commission.ctlGrid ctlGrid3 
         Height          =   2235
         Left            =   660
         TabIndex        =   27
         Top             =   6660
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3942
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "บันทึกรายการหักต่างๆ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   660
         TabIndex        =   28
         Top             =   6360
         Width           =   8835
      End
      Begin VB.Label lblProcess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "คำนวณการจ่ายค่าคอมฯ ของ ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   660
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   8835
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      TabIndex        =   13
      Top             =   -120
      Width           =   2655
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "คิดรวมยอดเบตเตอร์"
         Height          =   375
         Left            =   180
         TabIndex        =   33
         Top             =   6120
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "คิดรวมยอดอัคคีปราการ"
         Height          =   375
         Left            =   180
         TabIndex        =   32
         Top             =   6120
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   60
         TabIndex        =   29
         Top             =   6900
         Visible         =   0   'False
         Width           =   2475
         Begin VB.OptionButton OptJobType1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ค่าคอมฯปกติ"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptJobType2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ค่าคอมฯ งาน Job"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   540
            Width           =   1635
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ติ๊กถูก คิดรวมค่าขนส่ง (แบบเก่า)"
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   5700
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "เลือกประเภทการคำนวณ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1875
         Left            =   0
         TabIndex        =   20
         Top             =   3720
         Width           =   2655
         Begin VB.ComboBox cboSup 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1440
            Width           =   2235
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sup"
            Height          =   195
            Left            =   1440
            TabIndex        =   6
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sales"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtSale 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   180
            TabIndex        =   7
            Text            =   "*"
            Top             =   780
            Width           =   2235
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "หัวหน้าส่วนการขาย"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ผู้แทนขาย"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   540
            Width           =   735
         End
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   2235
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "ประมวลผล"
         Height          =   435
         Left            =   240
         TabIndex        =   10
         Top             =   7920
         Width           =   2115
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8580
         Width           =   2115
      End
      Begin Commission.ctlDate dtFrom 
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   1200
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate dtTo 
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   1800
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   120
         Top             =   7380
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin Commission.ctlDate ctlDate1 
         Height          =   375
         Left            =   195
         TabIndex        =   3
         Top             =   2640
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate ctlDate2 
         Height          =   375
         Left            =   195
         TabIndex        =   4
         Top             =   3240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   3000
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่รับชำระ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   18
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   16
         Top             =   1560
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   15
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmComPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_ As String
Dim curCustomer_ As BWGCustomer
Dim Counts As Long
Dim lvIndex As Long
Dim tList As ListItem

Dim tmpTable_Comm As String
Dim tmpTable_CommFee As String
Dim tmpTable_Comm_Re As String
Dim tmpTable_Comm_Sup As String
Dim tmpTable_Comm_Sup_Re As String

Dim tSql As String
Dim RS As ADODB.Recordset

Dim ProcessComplete As Boolean

Private Function GetComRate(ByVal PercentCom As Double, ByVal StaffType As Integer, Optional ByVal isJob As String = "N", Optional ByVal InPro As Integer = 0, Optional ByVal ForYearMonth As String) As Double
'On Error Resume Next
    GetComRate = 0
    If isJob = "N" Then
        If StaffType = 1 Then
            If UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "AKP" Then
                tSql = "SELECT     " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END)  AS PercentRate" & vbCrLf
                tSql = tSql & "From dbo.Com_tbComRateStaffAP" & vbCrLf
            ElseIf UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "PST" Then
                tSql = "SELECT     " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END)  AS PercentRate" & vbCrLf
                tSql = tSql & "From dbo.Com_tbComRateStaffPST" & vbCrLf
            Else
                tSql = "SELECT     " & PercentCom & " AS PercentCom, CASE WHEN " & InPro & " = 1 THEN CASE WHEN " & PercentCom & " >= 31 AND" & vbCrLf
                tSql = tSql & "                      " & PercentCom & " <= 50 THEN 1 ELSE MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END)" & vbCrLf
                tSql = tSql & "                      END ELSE MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END) END AS PercentRate" & vbCrLf
                tSql = tSql & "From dbo.Com_tbComRateStaff" & vbCrLf
            End If
            tSql = tSql & "WHERE CAST(ForYear AS Varchar(4)) + REPLACE(STR(CAST(ForMonth AS Varchar(2)), 2),' ','0') = '" & ForYearMonth & "'"
        ElseIf StaffType = 2 Then
            If UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "AKP" Then
                tSql = "SELECT " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END) PercentRate FROM Com_tbComRateSupAP"
            ElseIf UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "PST" Then
                tSql = "SELECT " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END) PercentRate FROM Com_tbComRateSupPST"
            Else
                tSql = "SELECT " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END) PercentRate FROM Com_tbComRateSup"
            End If
        End If
        Set RS = New ADODB.Recordset
        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        If Not RS.EOF Then GetComRate = IIf(IsNull(RS!PercentRate), 0, RS!PercentRate)
    Else
        If StaffType = 1 Then
            If UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "AKP" Then
                GetComRate = 0
            ElseIf UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "PST" Then
                GetComRate = 0
            Else
                GetComRate = 1.3
            End If
        ElseIf StaffType = 2 Then
            If UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "AKP" Then
                GetComRate = 0
            ElseIf UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "PST" Then
                GetComRate = 0
            Else
                GetComRate = 0.9
            End If
        End If
    End If
End Function

Private Function GetComRateExtra(ByVal PriceRate As Double, ByVal StaffType As Integer, ByVal WasteType As String, Optional ByVal ForYearMonth As String) As Double
On Error GoTo ErrD
    GetComRateExtra = 0
    If UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "AKP" Then
'        tSql = "SELECT     " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END)  AS PercentRate" & vbCrLf
'        tSql = tSql & "From dbo.Com_tbComRateStaffAP" & vbCrLf
        GetComRateExtra = 0
    ElseIf UCase(CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode) = "PST" Then
'        tSql = "SELECT     " & PercentCom & " AS PercentCom, MAX(CASE WHEN " & PercentCom & " >= PercentCom THEN PercentRate ELSE 0 END)  AS PercentRate" & vbCrLf
'        tSql = tSql & "From dbo.Com_tbComRateStaffPST" & vbCrLf
        GetComRateExtra = 0
    Else
        tSql = "SELECT '" & PriceRate & "' = PriceRate," & vbCrLf
        tSql = tSql & "CASE" & vbCrLf
        tSql = tSql & "    WHEN " & StaffType & " = 1 THEN CASE WHEN '" & PriceRate & "'  >= PriceRate THEN  PercentRateSale ELSE 0 END" & vbCrLf
        tSql = tSql & "    WHEN " & StaffType & " = 2 THEN CASE WHEN '" & PriceRate & "'  >= PriceRate THEN   PercentRateSup ELSE 0 END" & vbCrLf
        tSql = tSql & "    ELSE 0" & vbCrLf
        tSql = tSql & "END As PercentRate" & vbCrLf
        tSql = tSql & "From Com_tbComRateExpressBWG" & vbCrLf
        tSql = tSql & "WHERE CAST(ForYear AS Varchar(4)) + REPLACE(STR(CAST(ForMonth AS Varchar(2)), 2),' ','0') = '" & ForYearMonth & "' AND WasteType = '" & WasteType & "'"
        Set RS = New ADODB.Recordset
        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        If Not RS.EOF Then GetComRateExtra = RS!PercentRate
    
    End If
    Exit Function
    
ErrD:
    GetComRateExtra = 0
End Function

Private Sub ProcessCom()
Dim RS As ADODB.Recordset
Dim StaffName As String
Dim tSql As String
Dim i As Integer, ManifestFee As Double, ReValue As Double, Deduck As Double
Dim StaffType As Integer
    ProcessComplete = True
    If Option1.Value = True Then
        If Trim(txtSale.Text) = "*" Or Trim(txtSale.Text) = "" Then
            MsgBox "กรุณาเลือกเจ้าหน้าที่การตลาด", vbInformation, "Waring"
            lblProcess.Caption = "กรุณาเลือกเจ้าหน้าที่การตลาด"
            lblProcess.ForeColor = vbRed
            Exit Sub
        End If
        StaffName = Trim(txtSale.Text)
        lblProcess.Caption = "กำลังคำนวณค่าคอมฯ ของ ... " & txtSale.Text
        lblProcess.ForeColor = vbBlue
        DoEvents
        StaffType = 1
        If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
            Call ProcessComStaff(1, curStaff_.ID, txtSale.Text, 1, Check3.Value)
        Else
            Call ProcessComStaff(1, curStaff_.ID, txtSale.Text, 1, Check2.Value)
        End If
    ElseIf Option2.Value = True Then
        If cboSup.ListIndex = 0 Then
            MsgBox "กรุณาเลือกหัวหน้าส่วนการตลาด", vbInformation, "Warning"
            lblProcess.Caption = "กรุณาเลือกหัวหน้าส่วนการตลาด"
            lblProcess.ForeColor = vbRed
            Exit Sub
        Else
            StaffType = 2
            StaffName = cboSup.Text
            lblProcess.Caption = "กำลังคำนวณค่าคอมฯ ของ ... " & cboSup.Text
            lblProcess.ForeColor = vbBlue
            DoEvents
            
            tSql = "SELECT DISTINCT SaleStaffID, SaleStaffName FROM tbCom_SaleSup" & vbCrLf
            tSql = tSql & "WHERE SaleZoneID = '" & GetZoneID(cboSup) & "'" & vbCrLf
            tSql = tSql & "AND ForYear = " & dtFrom.YearValue & " AND ForMonth >= " & dtFrom.MonthValue & " AND ForMonth <= " & dtTo.MonthValue & vbCrLf
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            ProgressBar1.Min = 0
            ProgressBar1.Value = 0
            If RS.RecordCount > 0 Then
                ProgressBar1.Max = RS.RecordCount
            Else
                ProgressBar1.Max = 1
            End If
            Do While Not RS.EOF
                ProgressBar1.Max = RS.RecordCount
                ProgressBar1.Value = RS.AbsolutePosition
                DoEvents
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
                    Call ProcessComStaff(StaffType, RS!SaleStaffID, cboSup.Text, RS.AbsolutePosition, Check3.Value)
                Else
                    Call ProcessComStaff(StaffType, RS!SaleStaffID, cboSup.Text, RS.AbsolutePosition, Check2.Value)
                End If
                RS.MoveNext
            Loop
        End If
    End If
    If ProcessComplete = True Then
        If StaffType = 2 Then
            tSql = "SELECT * FROM " & tmpTable_Comm_Sup & " ORDER BY ForYear, ForMonth"
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            Do While Not RS.EOF
                For i = 1 To ctlGrid3.getGridObj.Rows - 1
                    If ctlGrid3.getGridObj.TextMatrix(i, 0) = RS!ForMonth Then
                        Deduck = IIf(IsNumeric(ctlGrid3.getGridObj.TextMatrix(i, 1)), ctlGrid3.getGridObj.TextMatrix(i, 1), 0)
                        Exit For
                    End If
                Next
                tSql = "UPDATE " & tmpTable_Comm_Sup & " SET InCollectFee = " & RS!InCollectFee - Deduck
                tSql = tSql & "WHERE CompanyID = '" & Trim(RS!CompanyID) & "'" & vbCrLf
                tSql = tSql & "AND ForYear = '" & Trim(RS!ForYear) & "' AND ForMonth = '" & Trim(RS!ForMonth) & "'"
                DBConnExc tSql
                
                For i = 1 To ctlGrid1.getGridObj.Rows - 1
                    If ctlGrid1.getGridObj.TextMatrix(i, 1) = RS!ForMonth Then
                        If IsNumeric(ctlGrid1.getGridObj.TextMatrix(i, 3)) = False Then
                            ManifestFee = 0
                        Else
                            ManifestFee = RS!InCollectFee - RS!TripTransFee_Inc - Deduck
                        End If
                        ctlGrid1.getGridObj.TextMatrix(i, 3) = Format(ManifestFee, "#,##0.00")
                        ctlGrid2.getGridObj.TextMatrix(i, 1) = Format(ManifestFee, "#,##0.00")
                        
                        ctlGrid1.getGridObj.TextMatrix(i, 4) = Format(CDbl(ctlGrid1.getGridObj.TextMatrix(i, 3)) * 100 / CDbl(ctlGrid1.getGridObj.TextMatrix(i, 2)), "0.00")
                        ctlGrid1.getGridObj.TextMatrix(i, 5) = Format(GetComRate(CDbl(ctlGrid1.getGridObj.TextMatrix(i, 4)), StaffType), "0.00")
                        ctlGrid2.getGridObj.TextMatrix(i, 3) = ctlGrid1.getGridObj.TextMatrix(i, 5)
                        Exit For
                    Else
                        ManifestFee = 0
                    End If
                Next
                
                RS.MoveNext
            Loop
            Set RS = Nothing
                
            tSql = "SELECT * FROM " & tmpTable_Comm_Sup_Re & " ORDER BY ForYear, ForMonth"
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            Do While Not RS.EOF
                For i = 1 To ctlGrid2.getGridObj.Rows - 1
                    If ctlGrid2.getGridObj.TextMatrix(i, 0) = RS!ForMonth Then
                    If IsNumeric(RS!ReValue) = False Then
                        ReValue = 0
                    Else
                        ReValue = RS!ReValue
                    End If
                    ctlGrid2.getGridObj.TextMatrix(i, 2) = Format(ReValue, "#,##0.00")
                    ctlGrid2.getGridObj.TextMatrix(i, 4) = Format(ReValue * CDbl(ctlGrid2.getGridObj.TextMatrix(i, 3)) / 100, "#,##0.00")
                    DoEvents
                    Exit For
                End If
                Next
                RS.MoveNext
            Loop
            RS.Close
            Set RS = Nothing
        End If
        
        cmdPrint.Visible = True
        lblProcess.Caption = "คำนวณการจ่ายค่าคอมฯ ของ ... " & StaffName & "  เสร็จแล้ว"
        lblProcess.ForeColor = vbMagenta
    Else
        cmdPrint.Visible = False
        lblProcess.Caption = "คำนวณการจ่ายค่าคอมฯ ของ ... " & StaffName & "  ล้มเหลว กรุณาลองใหม่"
        lblProcess.ForeColor = vbRed
    End If
End Sub

Private Sub ProcessComStaff(ByVal StaffType As Integer, ByVal StaffID As String, ByVal StaffName As String, Optional ByVal N As Integer = 1, Optional ByVal isCover As Integer = 0)
'On Error GoTo ErrD:
Dim i As Integer
Dim CompanyID As String
Dim ReValue As Double
Dim tmpSql As String
Dim RSInsert As ADODB.Recordset
Dim isJob As String, ExtraRate As Double
Dim ForYear As Integer, ForMonth As Integer, InPro As Integer

    If OptJobType2.Value = True Then
        isJob = "Y"
    Else
        isJob = "N"
    End If
    tmpTable_Comm_Re = "TEMPCommission_RE_" & ComName
    tmpTable_Comm_Sup = "tbCom_Commission_Sup_" & ComName
    tmpTable_Comm_Sup_Re = "tbCom_Commission_Sup_RE_" & ComName
    
    CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID
    
    If StaffType = 1 Then
        tSql = "SELECT  *  From Com_vwCommission_Target WHERE CompanyID  = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
        tSql = tSql & "AND SaleStaffID = '" & StaffID & "'" & vbCrLf
        tSql = tSql & "AND ForYear >= " & dtFrom.YearValue & " AND ForYear <= " & dtTo.YearValue & " AND ForMonth >= " & dtFrom.MonthValue & " AND ForMonth <= " & dtTo.MonthValue & vbCrLf
        tSql = tSql & "ORDER BY ForYear, ForMonth"
    Else
        tSql = "SELECT     CompanyID, SaleZoneID, ForYear, ForMonth, SUM(Target) AS Target, SUM(HICTarget) AS HICTarget, SUM(HBITarget) AS HBITarget, SUM(BBHTarget) AS BBHTarget, SUM(BBLTarget) AS BBLTarget, 0 AS InPro" & vbCrLf
        tSql = tSql & "From Com_vwCommission_Target_Sup " & vbCrLf
        tSql = tSql & "WHERE CompanyID  = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
        If GetZoneID(cboSup) <> "9716CB0E8EF94303A6CC09D03C8B4C25" Then tSql = tSql & "AND SaleZoneID = '" & GetZoneID(cboSup) & "'" & vbCrLf
        tSql = tSql & "AND ForYear >= " & dtFrom.YearValue & " AND ForYear <= " & dtTo.YearValue & " AND ForMonth >= " & dtFrom.MonthValue & " AND ForMonth <= " & dtTo.MonthValue & vbCrLf
        tSql = tSql & "GROUP BY CompanyID, SaleZoneID, ForYear, ForMonth"
    End If
    
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        ctlGrid1.getGridObj.Rows = RS.RecordCount + 1
        ctlGrid2.getGridObj.Rows = RS.RecordCount + 1
        ProgressBar1.Min = 0
        ProgressBar1.Value = 0
        If RS.RecordCount > 0 Then
            ProgressBar1.Max = RS.RecordCount
        Else
            ProgressBar1.Max = 1
        End If
        If isJob = "N" Then
            Do While Not RS.EOF
                ProgressBar1.Max = RS.RecordCount
                ProgressBar1.Value = RS.AbsolutePosition
                ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 0) = StaffName
                ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 1) = RS!ForYear & "" & Format(RS!ForMonth, "00")
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 2) = Format(RS!HICTarget + RS!HBITarget + RS!BBHTarget + RS!BBLTarget, "#,##0")
                Else
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 2) = Format(RS!Target, "#,##0")
                End If
                If RS!InPro = 0 Then
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 13) = "N"
                ElseIf RS!InPro = 1 Then
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 13) = "Y"
                End If
                RS.MoveNext
                DoEvents
            Loop
        End If
    End If
    RS.Close
    Set RS = Nothing
    
    If StaffType = 1 Then
        If Trim(txtSale.Text) <> "*" Or Trim(txtSale.Text) <> "" Then
           If VerifyData(StaffID, StaffType, , isJob) = False Then
                ProcessComplete = False
                MsgBox "เกิดข้อผิดพลาดขณะรวบรวมข้อมูล", vbInformation, "Waring"
                Exit Sub
           End If
        End If
    ElseIf StaffType = 2 Then
        If VerifyData(StaffID, StaffType, GetZoneID(cboSup), isJob) = False Then
             ProcessComplete = False
             MsgBox "เกิดข้อผิดพลาดขณะรวบรวมข้อมูล", vbInformation, "Waring"
             Exit Sub
        End If
    End If
    
    ExtraRate = 0
    For i = 1 To ctlGrid1.getGridObj.Rows - 1
        If ctlGrid1.getGridObj.TextMatrix(i, 3) = "" Then
            ctlGrid1.getGridObj.TextMatrix(i, 10) = Format(0, "0.00")
        ElseIf CDbl(ctlGrid1.getGridObj.TextMatrix(i, 3)) = 0 Then
            ctlGrid1.getGridObj.TextMatrix(i, 10) = Format(0, "0.00")
        ElseIf ctlGrid1.getGridObj.TextMatrix(i, 2) = "" Then
            ctlGrid1.getGridObj.TextMatrix(i, 10) = Format(0, "0.00")
        ElseIf CDbl(ctlGrid1.getGridObj.TextMatrix(i, 2)) = 0 Then
            ctlGrid1.getGridObj.TextMatrix(i, 10) = Format(0, "0.00")
        Else
            ctlGrid1.getGridObj.TextMatrix(i, 10) = Format(CDbl(ctlGrid1.getGridObj.TextMatrix(i, 3)) * 100 / CDbl(ctlGrid1.getGridObj.TextMatrix(i, 2)), "0.00")
        End If
'        ForYear = ctlDate1.YearValue
'        ForMonth = ctlGrid1.getGridObj.TextMatrix(i, 1)
        If ctlGrid1.getGridObj.TextMatrix(i, 13) = "Y" Then InPro = 1
        If ctlGrid1.getGridObj.TextMatrix(i, 13) = "N" Then InPro = 0
        If isJob = "N" Then
            ExtraRate = 0
            ctlGrid1.getGridObj.TextMatrix(i, 11) = Format(GetComRate(CDbl(ctlGrid1.getGridObj.TextMatrix(i, 10)), StaffType, , InPro, ctlGrid1.getGridObj.TextMatrix(i, 1)), "0.00")
            ExtraRate = ExtraRate + GetComRateExtra(CDbl(ctlGrid1.getGridObj.TextMatrix(i, 8)), StaffType, "Non", ctlGrid1.getGridObj.TextMatrix(i, 1))
            ExtraRate = ExtraRate + GetComRateExtra(CDbl(ctlGrid1.getGridObj.TextMatrix(i, 9)), StaffType, "Haz", ctlGrid1.getGridObj.TextMatrix(i, 1))
            ctlGrid1.getGridObj.TextMatrix(i, 12) = Format(ExtraRate, "0.00")
        Else
            ctlGrid1.getGridObj.TextMatrix(i, 10) = ""
            ctlGrid1.getGridObj.TextMatrix(i, 11) = Format(GetComRate(0, StaffType, isJob), "0.00")
            ctlGrid1.getGridObj.TextMatrix(i, 12) = "0.00"
        End If
        ctlGrid2.getGridObj.TextMatrix(i, 0) = ctlGrid1.getGridObj.TextMatrix(i, 1)
        ctlGrid2.getGridObj.TextMatrix(i, 1) = Format(ctlGrid1.getGridObj.TextMatrix(i, 3), "#,##0.00")
        ctlGrid2.getGridObj.TextMatrix(i, 3) = Format(ctlGrid1.getGridObj.TextMatrix(i, 11), "0.00")
        ctlGrid2.getGridObj.TextMatrix(i, 4) = Format(ctlGrid1.getGridObj.TextMatrix(i, 12), "0.00")
    Next

    'คำนวนยอดที่เก็บเงินได้ จาก Receive Vaucher
    tmpSql = "SELECT DISTINCT * " & vbCrLf
    tmpSql = tmpSql & "INTO " & tmpTable_Comm_Re & vbCrLf
    tmpSql = tmpSql & "FROM vw_ME_Bill_RE" & vbCrLf
    tmpSql = tmpSql & "WHERE WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "'"
    tmpSql = tmpSql & "AND REDate BETWEEN '" & ctlDate1.ValueYMD & "' AND '" & ctlDate2.ValueYMD & "'"
    tmpSql = tmpSql & "AND CompanyID  = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
    tmpSql = tmpSql & "AND UnderSaleStaffID = '" & StaffID & "'" & vbCrLf

    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_Comm_Re & "]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_Comm_Re & "]"
    DBConnExc tSql
    DBConnExc tmpSql
    
    tmpSql = "SELECT CompanyID, SaleDistrictID, UnderSaleStaffID, TimeTableNo, Year(WorkDate) AS ForYear, Month(WorkDate) AS ForMonth, SUM(InCollectFee) AS InCollectFee, MAX(TripTranFee) AS TripTransFee, MAX(TripTransFee_Inc) AS TripTransFee_Inc, MAX(VDTransFee) AS VDTransFee," & vbCrLf
    tmpSql = tmpSql & "CASE" & vbCrLf
    tmpSql = tmpSql & "    WHEN MAX(TripTranFee) > 0 THEN" & vbCrLf
    tmpSql = tmpSql & "                                    CASE WHEN MAX(TripTranFee) < MAX(VDTransFee) THEN  MAX(VDTransFee) - MAX(TripTranFee) ELSE 0 END" & vbCrLf
    tmpSql = tmpSql & "    WHEN MAX(TripTranFee) = 0 THEN 0" & vbCrLf
    tmpSql = tmpSql & "End As TripTransFeeDiff" & vbCrLf
    tmpSql = tmpSql & "INTO " & tmpTable_Comm_Re & "_New" & vbCrLf
    tmpSql = tmpSql & "FROM " & tmpTable_Comm_Re & vbCrLf
    tmpSql = tmpSql & "GROUP BY CompanyID, SaleDistrictID, UnderSaleStaffID, Year(WorkDate), Month(WorkDate), TimeTableNo"
    
    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_Comm_Re & "_New" & "]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_Comm_Re & "_New" & "]"
    DBConnExc tSql
    DBConnExc tmpSql
    
    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_Comm_Re & "]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_Comm_Re & "]"
    DBConnExc tSql
    
    
    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Com_tbCalcRE_SumTotal_All]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[Com_tbCalcRE_SumTotal_All]"
    DBConnExc tSql
    
    tSql = "SELECT CompanyID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth, SUM(InCollectFee) AS InCollectFee, SUM(TripTransFee) AS TripTransFee, SUM(TripTransFee_Inc) AS TripTransFee_Inc, SUM(VDTransFee) AS VDTransFee, SUM(TripTransFeeDiff) AS TripTransFeeDiff, 0.00 AS ExtraRate" & vbCrLf
    tSql = tSql & "INTO Com_tbCalcRE_SumTotal_All" & vbCrLf
    tSql = tSql & "FROM " & tmpTable_Comm_Re & "_New" & vbCrLf
    tSql = tSql & "GROUP BY CompanyID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth"
    DBConnExc tSql
    
    i = 0
    tSql = " SELECT * FROM Com_tbCalcRE_SumTotal_All"
    Set RS = New ADODB.Recordset
    RS.Open tSql, DBConn, adOpenKeyset, adLockOptimistic
    For i = 0 To RS.RecordCount - 1
        RS!ExtraRate = ctlGrid1.getGridObj.TextMatrix(i + 1, 12)
        RS.Update
        RS.MoveNext
    Next
    Set RS = Nothing
    
    tSql = "SELECT CompanyID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth, SUM(InCollectFee) AS InCollectFee, SUM(TripTransFee) AS TripTransFee, SUM(TripTransFee_Inc) AS TripTransFee_Inc, SUM(VDTransFee) AS VDTransFee, SUM(TripTransFeeDiff) AS TripTransFeeDiff" & vbCrLf
    tSql = tSql & "FROM " & tmpTable_Comm_Re & "_New" & vbCrLf
    tSql = tSql & "GROUP BY CompanyID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth"
    
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        ReValue = 0
        For i = 1 To ctlGrid2.getGridObj.Rows - 1
            ctlGrid2.getGridObj.TextMatrix(i, 2) = Format(0, "#,##0.00")
            ctlGrid2.getGridObj.TextMatrix(i, 5) = Format(0, "#,##0.00")
        Next
        ProgressBar1.Value = 0
        ProgressBar1.Min = 0
        If RS.RecordCount > 0 Then
            ProgressBar1.Max = RS.RecordCount
        Else
            ProgressBar1.Max = 1
        End If
        Do While Not RS.EOF
            ProgressBar1.Max = RS.RecordCount
            ProgressBar1.Value = RS.AbsolutePosition
            DoEvents
            If StaffType = 1 Then
                For i = 1 To ctlGrid2.getGridObj.Rows - 1
                    If ctlGrid2.getGridObj.TextMatrix(i, 0) = Format(RS!ForYear, "0000") & Format(RS!ForMonth, "00") Then
                        If IsNumeric(RS!InCollectFee) = False Then
                            ReValue = 0
                        Else
                            ReValue = RS!InCollectFee - RS!TripTransFee_Inc - RS!TripTransFeeDiff
'                            If RS!TripTransFee - RS!VDTransFee < 0 Then ReValue = ReValue - (RS!TripTransFee - RS!VDTransFee)
                        End If
                        ctlGrid2.getGridObj.TextMatrix(i, 2) = Format(ReValue, "#,##0.00")
                        ctlGrid2.getGridObj.TextMatrix(i, 5) = Format(ReValue * CDbl(ctlGrid2.getGridObj.TextMatrix(i, 3)) / 100, "#,##0.00")
                        Exit For
                    End If
                Next
            ElseIf StaffType = 2 Then
                For i = 1 To ctlGrid2.getGridObj.Rows - 1
                    If ctlGrid2.getGridObj.TextMatrix(i, 0) = Format(RS!ForYear, "0000") & Format(RS!ForMonth, "00") Then
                    If IsNumeric(ctlGrid2.getGridObj.TextMatrix(i, 2)) = False Then
                        ReValue = 0
                    Else
                        ReValue = CDbl(ctlGrid2.getGridObj.TextMatrix(i, 2))
                    End If
                    ReValue = ReValue + RS!InCollectFee - RS!TripTransFee_Inc - RS!TripTransFeeDiff
'                    If RS!TripTransFee - RS!VDTransFee < 0 Then ReValue = ReValue - (RS!TripTransFee - RS!VDTransFee)
                    ctlGrid2.getGridObj.TextMatrix(i, 2) = Format(ReValue, "#,##0.00")
                    ctlGrid2.getGridObj.TextMatrix(i, 5) = Format(ReValue * CDbl(ctlGrid2.getGridObj.TextMatrix(i, 3)) / 100, "#,##0.00")
                    DoEvents
                    Exit For
                End If
                Next
            End If
            
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS = Nothing
    
    If StaffType = 2 Then
        If N = 1 Then
            tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_Comm_Sup & "]') AND type in (N'U'))" & vbCrLf
            tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_Comm_Sup & "]" & vbCrLf
            DBConnExc tSql
            
            tSql = "SELECT * " & vbCrLf
            tSql = tSql & "INTO " & tmpTable_Comm_Sup & vbCrLf
            tSql = tSql & "From " & tmpTable_Comm
            DBConnExc tSql
        Else
            tSql = "SELECT * From " & tmpTable_Comm & " ORDER BY ForYear, ForMonth"
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            ProgressBar1.Value = 0
            ProgressBar1.Min = 0
            If RS.RecordCount > 0 Then
                ProgressBar1.Max = RS.RecordCount
            Else
                ProgressBar1.Max = 1
            End If
            Do While Not RS.EOF
                ProgressBar1.Max = RS.RecordCount
                ProgressBar1.Value = RS.AbsolutePosition
                DoEvents
                tSql = "SELECT * FROM " & tmpTable_Comm_Sup & vbCrLf
                tSql = tSql & "WHERE CompanyID = '" & Trim(RS!CompanyID) & "'" & vbCrLf
                tSql = tSql & "AND ForYear = '" & Trim(RS!ForYear) & "' AND ForMonth = '" & Trim(RS!ForMonth) & "'"
                Set RSInsert = New ADODB.Recordset
                GetDBConnection
                RSInsert.Open tSql, DBConn, adOpenKeyset, adLockOptimistic
                If RSInsert.EOF Then
                    RSInsert.AddNew
                    RSInsert!CompanyID = RS!CompanyID
                    RSInsert!SaleZoneID = RS!SaleZoneID
'                    RSInsert!UnderSaleStaffID = RS!UnderSaleStaffID
                    RSInsert!Target = RS!Target
                    RSInsert!ForMonth = RS!ForMonth
                    RSInsert!ForYear = RS!ForYear
                    RSInsert!InCollectFee = RS!InCollectFee
                    RSInsert!TripTransFee = RS!TripTransFee
                    RSInsert!NonFee = RS!NonFee
                    RSInsert!HazFee = RS!HazFee
                    RSInsert!NBLFee = RS!NBLFee
                    RSInsert!HBLFee = RS!HBLFee
                    RSInsert!HICFee = RS!HICFee
                    RSInsert!HBIFee = RS!HBIFee
                    RSInsert!HICFee = RS!BBHFee
                    RSInsert!HBIFee = RS!BBLFee
                    RSInsert!HICQty = RS!HICQty
                    RSInsert!HBIQty = RS!HBIQty
                    RSInsert!HICQty = RS!BBHQty
                    RSInsert!HBIQty = RS!BBLQty
                    RSInsert!TripTransFee_Inc = RS!TripTransFee_Inc
                    RSInsert!TimeTableNo = ""
                    RSInsert!MinWorkDate = ""
                    RSInsert!MaxWorkDate = ""
                    RSInsert!docNo = ""
                Else
                    RSInsert!CompanyID = RS!CompanyID
                    RSInsert!SaleZoneID = RS!SaleZoneID
                    RSInsert!Target = RSInsert!Target + RS!Target
                    RSInsert!ForMonth = RS!ForMonth
                    RSInsert!ForYear = RS!ForYear
                    RSInsert!InCollectFee = RSInsert!InCollectFee + RS!InCollectFee
                    RSInsert!TripTransFee = RSInsert!TripTransFee + RS!TripTransFee
                    RSInsert!NonFee = RSInsert!NonFee + RS!NonFee
                    RSInsert!HazFee = RSInsert!HazFee + RS!HazFee
                    RSInsert!NBLFee = RSInsert!NBLFee + RS!NBLFee
                    RSInsert!HBLFee = RSInsert!HBLFee + RS!HBLFee
                    RSInsert!HICFee = RSInsert!HICFee + RS!HICFee
                    RSInsert!HBIFee = RSInsert!HBIFee + RS!HBIFee
                    RSInsert!BBHFee = RSInsert!BBHFee + RS!BBHFee
                    RSInsert!BBLFee = RSInsert!BBLFee + RS!BBLFee
                    RSInsert!HICQty = RSInsert!HICQty + RS!HICQty
                    RSInsert!HBIQty = RSInsert!HBIQty + RS!HBIQty
                    RSInsert!BBHQty = RSInsert!BBHQty + RS!BBHQty
                    RSInsert!BBLQty = RSInsert!BBLQty + RS!BBLQty
                    RSInsert!TripTransFee_Inc = RSInsert!TripTransFee_Inc + RS!TripTransFee_Inc
                    RSInsert!TimeTableNo = ""
                End If
                RSInsert.Update
                RS.MoveNext
            Loop
            CloseDBConnection
        End If
    End If
    
    If StaffType = 2 Then
        If N = 1 Then
            tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_Comm_Sup_Re & "]') AND type in (N'U'))" & vbCrLf
            tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_Comm_Sup_Re & "]" & vbCrLf
            DBConnExc tSql
            
            If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID <> "776EDA5830F744A6B3551348D851BC22" Then
                tSql = "SELECT DISTINCT RE.* " & vbCrLf
                tSql = tSql & "INTO " & tmpTable_Comm_Sup_Re & vbCrLf
                tSql = tSql & "FROM         dbo.Com_tbCalcRE_SumTotal_All AS RE INNER JOIN" & vbCrLf
                tSql = tSql & "                      dbo.tbCom_SaleSup AS SS ON RE.CompanyID = SS.CompanyID AND RE.SaleZoneID = SS.SaleZoneID AND RE.SaleStaffID = SS.SaleStaffID AND" & vbCrLf
                tSql = tSql & "                      RE.ForYear = SS.ForYear And RE.ForMonth = SS.ForMonth" & vbCrLf
                tSql = tSql & "WHERE (SS.SaleStaffID = '" & StaffID & "')"
                tSql = tSql & "AND     (RE.SaleZoneID = '" & GetZoneID(cboSup) & "') " & vbCrLf
            Else
                tSql = "SELECT DISTINCT RE.* " & vbCrLf
                tSql = tSql & "INTO " & tmpTable_Comm_Sup_Re & vbCrLf
                tSql = tSql & "FROM         dbo.Com_tbCalcRE_SumTotal_All AS RE INNER JOIN" & vbCrLf
                tSql = tSql & "                      dbo.tbCom_SaleSup AS SS ON RE.SaleZoneID = SS.SaleZoneID AND RE.SaleStaffID = SS.SaleStaffID AND" & vbCrLf
                tSql = tSql & "                      RE.ForYear = SS.ForYear And RE.ForMonth = SS.ForMonth" & vbCrLf
                tSql = tSql & "WHERE (SS.SaleStaffID = '" & StaffID & "')"
            End If
            DBConnExc tSql
        Else
            If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID <> "776EDA5830F744A6B3551348D851BC22" Then
                tSql = "SELECT DISTINCT RE.* " & vbCrLf
                tSql = tSql & "FROM         dbo.Com_tbCalcRE_SumTotal_All AS RE INNER JOIN" & vbCrLf
                tSql = tSql & "                      dbo.tbCom_SaleSup AS SS ON RE.CompanyID = SS.CompanyID AND RE.SaleZoneID = SS.SaleZoneID AND RE.SaleStaffID = SS.SaleStaffID AND" & vbCrLf
                tSql = tSql & "                      RE.ForYear = SS.ForYear And RE.ForMonth = SS.ForMonth" & vbCrLf
                tSql = tSql & "WHERE (SS.SaleStaffID = '" & StaffID & "')"
                tSql = tSql & "AND     (RE.SaleZoneID = '" & GetZoneID(cboSup) & "') " & vbCrLf
            Else
                tSql = "SELECT DISTINCT RE.* " & vbCrLf
                tSql = tSql & "FROM         dbo.Com_tbCalcRE_SumTotal_All AS RE INNER JOIN" & vbCrLf
                tSql = tSql & "                      dbo.tbCom_SaleSup AS SS ON RE.SaleZoneID = SS.SaleZoneID AND RE.SaleStaffID = SS.SaleStaffID AND" & vbCrLf
                tSql = tSql & "                      RE.ForYear = SS.ForYear And RE.ForMonth = SS.ForMonth" & vbCrLf
                tSql = tSql & "WHERE (SS.SaleStaffID = '" & StaffID & "')"
            End If
           Set RS = New ADODB.Recordset
           Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            ProgressBar1.Value = 0
            ProgressBar1.Min = 0
            If RS.RecordCount > 0 Then
                ProgressBar1.Max = RS.RecordCount
            Else
                ProgressBar1.Max = 1
            End If
            
            Do While Not RS.EOF
                ProgressBar1.Max = RS.RecordCount
                ProgressBar1.Value = RS.AbsolutePosition
                DoEvents
                tSql = "SELECT * FROM " & tmpTable_Comm_Sup_Re & vbCrLf
                tSql = tSql & "WHERE ForYear = '" & Trim(RS!ForYear) & "' AND ForMonth = '" & Trim(RS!ForMonth) & "'"
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID <> "776EDA5830F744A6B3551348D851BC22" Then
                    tSql = tSql & "AND CompanyID = '" & Trim(RS!CompanyID) & "'" & vbCrLf
                End If
                Set RSInsert = New ADODB.Recordset
                GetDBConnection
                RSInsert.Open tSql, DBConn, adOpenKeyset, adLockOptimistic
                If RSInsert.EOF Then
                    RSInsert.AddNew
                    RSInsert!CompanyID = RS!CompanyID
                    RSInsert!SaleZoneID = RS!SaleZoneID
                    RSInsert!SaleStaffID = RS!SaleZoneID
                    RSInsert!StaffName = StaffName
                    RSInsert!ForMonth = RS!ForMonth
                    RSInsert!ForYear = RS!ForYear
                    RSInsert!ForYear = RS!ForYear
                    RSInsert!ManifestFee = RS!ManifestFee
                    RSInsert!TransFee = RS!TransFee
                    RSInsert!ExcTransFee = RS!ExcTransFee
                    RSInsert!ReValue = RS!ReValue
                    RSInsert.Update
                Else
                    RSInsert!CompanyID = RS!CompanyID
                    RSInsert!SaleZoneID = RS!SaleZoneID
                    RSInsert!SaleStaffID = RS!SaleZoneID
                    RSInsert!StaffName = StaffName
                    RSInsert!ForYear = RS!ForYear
                    RSInsert!ForMonth = RS!ForMonth
                    RSInsert!ManifestFee = RSInsert!ManifestFee + RS!ManifestFee
                    RSInsert!TransFee = RSInsert!TransFee + RS!TransFee
                    RSInsert!ExcTransFee = RSInsert!ExcTransFee + RS!ExcTransFee
                    RSInsert!ReValue = RSInsert!ReValue + RS!ReValue
                    RSInsert.Update
                End If
                RS.MoveNext
           Loop
           RS.Close
           Set RS = Nothing
            CloseDBConnection
        End If
    End If
    
    Exit Sub
ErrD:
    ProcessComplete = False
    Resume Next
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex <> 3 Then
        Check2.Visible = True
        Check3.Visible = False
    Else
        Check2.Visible = False
        Check3.Visible = True
    End If
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'On Error Resume Next
Dim SubRpt As String
    If ProcessComplete = False Then
        MsgBox "ไม่สามารถพิมพ์รายงานได้", vbInformation, "Warning"
        Exit Sub
    End If
    
    If Option1.Value = True Then
        tSql = "ALTER VIEW Com_vwComPay" & vbCrLf
        tSql = tSql & "AS" & vbCrLf
'        tSql = tSql & "SELECT     C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target, MAX(ISNULL(C.InCollectFee, 0)) - MAX(ISNULL(C.TripTransFee_Inc, 0))" & vbCrLf
'        tSql = tSql & "                      AS Income, SUM(ISNULL(C.TripTransFee_Inc, 0)) AS TripTransFee_Inc, SUM(ISNULL(RE.ManifestFee, 0)) AS ManifestFee, SUM(ISNULL(RE.TransFee, 0))" & vbCrLf
'        tSql = tSql & "                      AS TransFee, SUM(ISNULL(RE.ExcTransFee, 0)) AS ExcTransFee, SUM(ISNULL(RE.ManifestFee, 0)) - SUM(ISNULL(RE.TransFee, 0)) AS REValue" & vbCrLf
'        tSql = tSql & "FROM         dbo.Com_tbCalcRE_SumTotal_All AS RE RIGHT OUTER JOIN" & vbCrLf
'        tSql = tSql & "                      " & tmpTable_Comm & " AS C ON RE.SaleStaffID = C.SaleStaffID AND RE.ForYear = C.ForYear AND RE.ForMonth = C.ForMonth AND" & vbCrLf
'        tSql = tSql & "                      RE.SaleZoneID = C.SaleZoneID" & vbCrLf
'        tSql = tSql & "GROUP BY C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target
        
        tSql = tSql & "SELECT     C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target, MAX(ISNULL(C.InCollectFee, 0)) - MAX(ISNULL(C.TripTransFee_Inc, 0))" & vbCrLf
        tSql = tSql & "                      AS Income, SUM(ISNULL(C.TripTransFee_Inc, 0)) AS TripTransFee_Inc, SUM(ISNULL(RE.InCollectFee, 0)) AS ManifestFee," & vbCrLf
        tSql = tSql & "                      SUM(ISNULL(RE.TripTransFee_Inc, 0)) AS TransFee, SUM(ISNULL(RE.TripTransFee, 0)) AS ExcTransFee, SUM(ISNULL(RE.InCollectFee, 0))" & vbCrLf
        tSql = tSql & "                      - SUM(ISNULL(RE.TripTransFee_Inc, 0)) AS REValue, MAX(RE.ExtraRate) AS ExtraRate" & vbCrLf
        tSql = tSql & "FROM         Com_tbCalcRE_SumTotal_All AS RE RIGHT OUTER JOIN" & vbCrLf
        tSql = tSql & "                      " & tmpTable_Comm & " AS C ON RE.UnderSaleStaffID = C.SaleStaffID AND RE.SaleDistrictID = C.SaleDistrictID AND" & vbCrLf
        tSql = tSql & "                      RE.CompanyID = C.CompanyID And RE.ForYear = C.ForYear And RE.ForMonth = C.ForMonth" & vbCrLf
        tSql = tSql & "GROUP BY C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target"
        DBConnExc tSql
        
        tSql = "ALTER VIEW Com_vwComPay_Report" & vbCrLf
        tSql = tSql & "AS" & vbCrLf
        tSql = tSql & "SELECT * " & vbCrLf
        tSql = tSql & ", PercentCom = ([Income] / [Target]) * 100" & vbCrLf
        tSql = tSql & ", PercentRate = (SELECT Max(S.PercentRate)" & vbCrLf
        tSql = tSql & "                 FROM Com_tbComRateStaff S WHERE S.ForYear = C.ForYear AND S.ForMonth = C.ForMonth AND (([Income] / [Target]) * 100) >= S.PercentCom)" & vbCrLf
        tSql = tSql & "FROM Com_vwComPay C"
        DBConnExc tSql
        
    Else
        tSql = "ALTER VIEW Com_vwComPay" & vbCrLf
        tSql = tSql & "AS" & vbCrLf
        If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID <> "776EDA5830F744A6B3551348D851BC22" Then
            tSql = tSql & "SELECT     C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target, MAX(ISNULL(C.InCollectFee, 0)) - MAX(ISNULL(C.TripTransFee_Inc, 0))" & vbCrLf
            tSql = tSql & "                      AS Income, SUM(ISNULL(C.TripTransFee_Inc, 0)) AS TripTransFee_Inc, SUM(ISNULL(RE.ManifestFee, 0)) AS ManifestFee, SUM(ISNULL(RE.TransFee, 0))" & vbCrLf
            tSql = tSql & "                      AS TransFee, SUM(ISNULL(RE.ExcTransFee, 0)) AS ExcTransFee, SUM(ISNULL(RE.ManifestFee, 0)) - SUM(ISNULL(RE.TransFee, 0)) AS REValue" & vbCrLf
            tSql = tSql & "FROM         " & tmpTable_Comm_Sup_Re & " AS RE RIGHT OUTER JOIN" & vbCrLf
            tSql = tSql & "                      " & tmpTable_Comm_Sup & " AS C ON RE.SaleZoneID = C.SaleZoneID AND RE.ForYear = C.ForYear AND RE.ForMonth = C.ForMonth" & vbCrLf
            tSql = tSql & "GROUP BY C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target"
        Else
            tSql = tSql & "SELECT     C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target, MAX(ISNULL(C.InCollectFee, 0)) - MAX(ISNULL(C.TripTransFee_Inc, 0))" & vbCrLf
            tSql = tSql & "                      AS Income, SUM(ISNULL(C.TripTransFee_Inc, 0)) AS TripTransFee_Inc, SUM(ISNULL(RE.ManifestFee, 0)) AS ManifestFee, SUM(ISNULL(RE.TransFee, 0))" & vbCrLf
            tSql = tSql & "                      AS TransFee, SUM(ISNULL(RE.ExcTransFee, 0)) AS ExcTransFee, SUM(ISNULL(RE.ManifestFee, 0)) - SUM(ISNULL(RE.TransFee, 0)) AS REValue" & vbCrLf
            tSql = tSql & "FROM         " & tmpTable_Comm_Sup_Re & " AS RE RIGHT OUTER JOIN" & vbCrLf
            tSql = tSql & "                      " & tmpTable_Comm_Sup & " AS C ON RE.SaleZoneID = C.SaleZoneID AND RE.ForYear = C.ForYear AND RE.ForMonth = C.ForMonth" & vbCrLf
            tSql = tSql & "GROUP BY C.SaleZoneID, C.SaleStaffID, C.StaffFName, C.ForYear, C.ForMonth, C.Target"
        End If
        DBConnExc tSql
        
        tSql = "ALTER VIEW Com_vwComPay_Report AS " & vbCrLf
        tSql = tSql & "SELECT *  FROM Com_vwComPay"
        DBConnExc tSql
        
    End If
    
    With rpt1
        .Reset
        .WindowTitle = "รายงานการจ่ายค่าคอมฯ"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If Option1.Value = True Then
            .ReportFileName = App.Path & "\Reports\RptCompay_Sale.rpt"
        ElseIf Option2.Value = True Then
            If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
                .ReportFileName = App.Path & "\Reports\RptCompay_Sup_AP.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\RptCompay_Sup.rpt"
            End If
        End If
'        .Formulas(0) = "Company='บริษัท " & Trim(cboCompany.Text) & "'"
'        .Formulas(1) = "WorkDate='" & "วันที่เก็บเงิน  " & ctlDate1.ValueDMY & "  ถึงวันที่  " & ctlDate2.ValueDMY & "'"
        
        If Check1.Value = 0 Then
            .Formulas(3) = "IsIncTrans='N'"
        Else
            .Formulas(3) = "IsIncTrans='Y'"
        End If
        If Option1.Value = True Then
            .Formulas(2) = "SaleName='" & IIf(Trim(txtSale.Text) <> "", IIf(Trim(txtSale) <> "*", txtSale.Text, ""), "") & "'"
            .Formulas(4) = "IsSup = 'N'"
        Else
            .Formulas(2) = "SaleName='" & cboSup.Text & "'"
            .Formulas(4) = "IsSup = 'Y'"
        End If
        
        .SqlQuery = "SELECT * FROM Com_vwComPay_Report"
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
                      
End Sub

Private Sub cmdProcess_Click()
    lblProcess.Visible = True
    ProgressBar1.Visible = True
    ctlGrid1.Visible = True
    ctlGrid2.Visible = True
    ctlGrid3.Visible = True
    ctlGrid1.getGridObj.Rows = 1
    ctlGrid2.getGridObj.Rows = 1
    cmdPrint.Visible = False
    Call ProcessCom
End Sub

Private Sub ctlGrid3_UpdateValue(row As Integer, col As Integer)
    ctlGrid3.getGridObj.TextMatrix(row, col) = Format(ctlGrid3.getGridObj.TextMatrix(row, col), "#,##0.00")
End Sub

Private Sub dtFrom_Validate(Cancel As Boolean)
On Error Resume Next
Dim i As Integer, j As Integer
        ctlGrid3.getGridObj.Rows = dtTo.MonthValue - dtFrom.MonthValue + 2
        j = 0
        For i = dtFrom.MonthValue To dtTo.MonthValue
            j = j + 1
            ctlGrid3.getGridObj.TextMatrix(j, 0) = i
        Next
End Sub


Private Sub dtTo_Validate(Cancel As Boolean)
Dim i As Integer, j As Integer
        ctlGrid3.getGridObj.Rows = dtTo.MonthValue - dtFrom.MonthValue + 2
        j = 0
        For i = dtFrom.MonthValue To dtTo.MonthValue
            j = j + 1
            ctlGrid3.getGridObj.TextMatrix(j, 0) = i
        Next
End Sub


Private Sub Form_Activate()
    ActivateForm
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    ActivateForm
    PopulateCompany
    GetSup
    Option1_Click
    
    Check2.Visible = False
    Check3.Visible = False
    
End Sub

Private Sub ActivateForm()
    Frame1.Height = Me.Height - 350
    Frame2.Height = Me.Height - 350
    Frame2.Width = Me.Width - Frame2.Left - 150
    dtFrom.DayValue = "01"
    ctlDate1.DayValue = "01"
    
    lblProcess.Width = Me.Width - Frame2.Left - Frame1.Width
    Label3.Width = Me.Width - Frame2.Left - Frame1.Width
    ProgressBar1.Width = Me.Width - Frame2.Left - Frame1.Width
    ctlGrid1.Width = Me.Width - Frame1.Width - (Frame2.Left / 2)
    ctlGrid2.Width = Me.Width - Frame1.Width - (Frame2.Left / 2)
    ctlGrid3.Width = Me.Width - Frame1.Width - (Frame2.Left / 2)
    
    lblProcess.Left = (Frame2.Left) / 4
    Label3.Left = (Frame2.Left) / 4
    ProgressBar1.Left = (Frame2.Left) / 4
    ctlGrid1.Left = (Frame2.Left) / 4
    ctlGrid2.Left = (Frame2.Left) / 4
    ctlGrid3.Left = (Frame2.Left) / 4
    cmdPrint.Left = (Frame2.Left) / 4
    
    ctlGrid1.getGridObj.Cols = 14
    ctlGrid1.getGridObj.TextMatrix(0, 0) = "ผู้แทนขาย"
    ctlGrid1.getGridObj.TextMatrix(0, 1) = "เดือน"
    ctlGrid1.getGridObj.TextMatrix(0, 2) = "เป้าการขาย"
    ctlGrid1.getGridObj.TextMatrix(0, 3) = "ยอดการขาย"
    ctlGrid1.getGridObj.TextMatrix(0, 4) = "ยอด Non"
    ctlGrid1.getGridObj.TextMatrix(0, 5) = "ยอด Haz"
    ctlGrid1.getGridObj.TextMatrix(0, 6) = "W Non"
    ctlGrid1.getGridObj.TextMatrix(0, 7) = "W Haz"
    ctlGrid1.getGridObj.TextMatrix(0, 8) = "Avg Non"
    ctlGrid1.getGridObj.TextMatrix(0, 9) = "Avg Haz"
    ctlGrid1.getGridObj.TextMatrix(0, 10) = "% การขาย"
    ctlGrid1.getGridObj.TextMatrix(0, 11) = "% การจ่าย"
    ctlGrid1.getGridObj.TextMatrix(0, 12) = "% เพิ่มพิเศษ"
    ctlGrid1.getGridObj.TextMatrix(0, 13) = "ทดลองงาน"
    ctlGrid1.getGridObj.ColWidth(0) = 1200
    ctlGrid1.getGridObj.ColWidth(1) = 800
    ctlGrid1.getGridObj.ColWidth(2) = 1000
    ctlGrid1.getGridObj.ColWidth(3) = 1100
    ctlGrid1.getGridObj.ColWidth(4) = 1000
    ctlGrid1.getGridObj.ColWidth(5) = 1000
    ctlGrid1.getGridObj.ColWidth(6) = 800
    ctlGrid1.getGridObj.ColWidth(7) = 800
    ctlGrid1.getGridObj.ColWidth(8) = 800
    ctlGrid1.getGridObj.ColWidth(9) = 800
    ctlGrid1.getGridObj.ColWidth(10) = 1000
    ctlGrid1.getGridObj.ColWidth(11) = 1000
    ctlGrid1.getGridObj.ColWidth(12) = 1000
    ctlGrid1.getGridObj.ColWidth(13) = 900
    
    ctlGrid2.getGridObj.Cols = 6
    ctlGrid2.getGridObj.TextMatrix(0, 0) = "เดือน"
    ctlGrid2.getGridObj.TextMatrix(0, 1) = "ยอดการขาย"
    ctlGrid2.getGridObj.TextMatrix(0, 2) = "ยอดที่เรียกเก็บได้"
    ctlGrid2.getGridObj.TextMatrix(0, 3) = "อัตราส่วนการจ่าย"
    ctlGrid2.getGridObj.TextMatrix(0, 4) = "อัตราส่วนเพิ่มพิเศษ"
    ctlGrid2.getGridObj.TextMatrix(0, 5) = "ค่าคอมมิสชั่น"
    ctlGrid2.getGridObj.ColWidth(0) = 1500
    ctlGrid2.getGridObj.ColWidth(1) = 1500
    ctlGrid2.getGridObj.ColWidth(2) = 1500
    ctlGrid2.getGridObj.ColWidth(3) = 1500
    ctlGrid2.getGridObj.ColWidth(4) = 1500
    ctlGrid2.getGridObj.ColWidth(5) = 1500
    ctlGrid2.getGridObj.ColAlignment(0) = 4
    
    ctlGrid3.getGridObj.Cols = 2
    ctlGrid3.getGridObj.TextMatrix(0, 0) = "เดือน"
    ctlGrid3.getGridObj.TextMatrix(0, 1) = "หัก"
    ctlGrid3.getGridObj.ColWidth(0) = 1200
    ctlGrid3.getGridObj.ColWidth(1) = 1200
    ctlGrid3.getGridObj.ColAlignment(0) = 4
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

Private Sub GetSup()
On Error Resume Next
Dim x As Integer
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "SELECT SupStaffName" & vbCrLf
    tSql = tSql & "From dbo.tbCom_SaleSup" & vbCrLf
    tSql = tSql & "WHERE     (ForYear = '2551') AND (ForMonth = 6)" & vbCrLf
    tSql = tSql & "GROUP BY ZoneName, SupStaffName"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If RS.EOF Then Exit Sub
    cboSup.Clear
    cboSup.AddItem "เลือกหัวหน้าส่วนการขาย"
    Do While Not RS.EOF
        cboSup.AddItem RS!SupStaffName
        RS.MoveNext
    Loop
    Set RS = Nothing
    If cboSup.ListCount > 0 Then cboSup.ListIndex = 0
End Sub

Private Function GetZoneID(ByVal SupName) As String
Dim tSql As String
Dim RS As ADODB.Recordset
    GetZoneID = ""
    tSql = "SELECT SaleZoneID FROM tbCom_SaleSup WHERE SupStaffName = '" & Trim(SupName) & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then GetZoneID = RS!SaleZoneID
    RS.Close
    Set RS = Nothing
End Function

Private Sub Option1_Click()
    If Option1.Value = True Then
        cboSup.Enabled = False
        cboSup.ListIndex = 0
        txtSale.Enabled = True
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        txtSale.Enabled = False
        cboSup.Enabled = True
    End If
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If txtSale.Text = "" Then Exit Sub
    If txtSale.Text = "*" Then Exit Sub
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtSale.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Private Function VerifyData(ByVal StaffID As String, StaffType As Integer, Optional ByVal SaleZoneID As String = "", Optional ByVal isJob As String = "N") As Boolean
'On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim RSInsert As ADODB.Recordset
Dim tmpSql As String, tSql As String
Dim i As Integer, ManifestFee As Double, Deduck As Double

    tmpTable_Comm = "tbCom_Commission_" & ComName
    tmpTable_CommFee = "tbCom_Commission_" & ComName & "_Fee"
    
    VerifyData = True
    tmpSql = ""
    tmpSql = "SELECT     CompanyID, SaleStaffID, SaleDistrictID, ForYear, ForMonth, " & vbCrLf
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode = "AKP" Then
        tmpSql = tmpSql & "Target = HICTarget + HBITarget + BBHTarget + BBLTarget, " & vbCrLf
    Else
        tmpSql = tmpSql & "Target, "
    End If
    tmpSql = tmpSql & "CONVERT(Numeric(18,2), 0) AS InCollectFee, CONVERT(Numeric(18,2), 0) AS  TripTransFee, CONVERT(Numeric(18,2), 0) AS TripTransFee_Inc, CONVERT(Numeric(18,2), 0) AS VDTransFee, CONVERT(Numeric(18,2), 0) AS TripTransFeeDiff, NonTarget, CONVERT(Numeric(18,2), 0) AS  NonFee, HazTarget, CONVERT(Numeric(18,2), 0) AS HazFee, NBLTarget, CONVERT(Numeric(18,2), 0) AS  NBLFee, " & vbCrLf
    tmpSql = tmpSql & "HBLTarget, CONVERT(Numeric(18,2), 0) AS  HBLFee, HICTarget, CONVERT(Numeric(18,2), 0) AS  HICFee, HBITarget, CONVERT(Numeric(18,2), 0) AS  HBIFee, BBHTarget, CONVERT(Numeric(18,2), 0) AS  BBHFee, CONVERT(Numeric(18,2), 0) AS  BBLFee, BBLTarget, StaffFName, StaffLName, '' TimeTableNo, SaleZoneID, SDName, ZoneName, '' MinWorkDate," & vbCrLf
    tmpSql = tmpSql & "'' MaxWorkDate, '' DocNo, CONVERT(Numeric(18,2), 0) AS NonQty, CONVERT(Numeric(18,2), 0) AS NBLQty, CONVERT(Numeric(18,2), 0) AS HazQty, CONVERT(Numeric(18,2), 0) AS HBLQty, CONVERT(Numeric(18,2), 0) AS HICQty, CONVERT(Numeric(18,2), 0) AS HBIQty, CONVERT(Numeric(18,2), 0) AS BBHQty, CONVERT(Numeric(18,2), 0) AS BBLQty" & vbCrLf
    tmpSql = tmpSql & "INTO " & tmpTable_Comm & vbCrLf
    tmpSql = tmpSql & "From Com_vwCommission_Target" & vbCrLf
    tmpSql = tmpSql & "WHERE (ForYear >= '" & dtFrom.YearValue & "'"
    tmpSql = tmpSql & " AND ForYear <= '" & dtTo.YearValue & "'"
    tmpSql = tmpSql & " AND ForMonth >= '" & dtFrom.MonthValue & "'"
    tmpSql = tmpSql & " AND ForMonth <= '" & dtTo.MonthValue & "')" & vbCrLf
    
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode = "AKP" Then
'        tmpSql = tmpSql & " AND (HICTarget > 0 OR HBITarget > 0)"
        tmpSql = tmpSql & " And (CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
    Else
        tmpSql = tmpSql & " And (CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
    End If

    tmpSql = tmpSql & "AND SaleStaffID = '" & StaffID & "'" & vbCrLf
    If StaffType = 2 Then
        If SaleZoneID <> "9716CB0E8EF94303A6CC09D03C8B4C25" Then tmpSql = tmpSql & "AND SaleZoneID = '" & SaleZoneID & "'" & vbCrLf
    End If
        

    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_Comm & "]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_Comm & "]"
    DBConnExc tSql
    DBConnExc tmpSql
    DoEvents
    'คำนวนยอดขายจากใบเมนิเฟส
    tmpSql = ""
    tmpSql = "SELECT     CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForMonth, ForYear, SUM(InCollectFee) AS InCollectFee, CONVERT(Numeric(18,2), 0) AS TripTransFee, CONVERT(Numeric(18,2), 0) AS TripTransFee_Inc, CONVERT(Numeric(18,2), 0) AS VDTransFee, CONVERT(Numeric(18,2), 0) AS TripTransFeeDiff, SUM(NonFee) AS NonFee, " & vbCrLf
    tmpSql = tmpSql & "SUM(HazFee) AS HazFee, SUM(NBLFee) AS NBLFee, SUM(HBLFee) AS HBLFee, SUM(HICFee) AS HICFee, SUM(HBIFee) AS HBIFee, SUM(BBHFee) AS BBHFee, SUM(BBLFee) AS BBLFee, SUM(NonQty) AS NonQty, SUM(NBLQty) AS NBLQty, SUM(HazQty) AS HazQty, SUM(HBLQty) AS HBLQty, SUM(HICQty) AS HICQty, SUM(HBIQty) AS HBIQty, SUM(BBHQty) AS BBHQty, SUM(BBLQty) AS BBLQty, 0 AS ExtraRate" & vbCrLf
    tmpSql = tmpSql & "INTO " & tmpTable_CommFee & vbCrLf
    tmpSql = tmpSql & "From dbo.Com_vwCommission_Fee C" & vbCrLf
    tmpSql = tmpSql & "Where MinWorkDate >= '" & dtFrom.ValueYMD & "'"
    tmpSql = tmpSql & " AND MaxWorkDate <= '" & dtTo.ValueYMD & "'" & vbCrLf
    
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode = "AKP" Then
        If Check3.Value = 0 Then
            tmpSql = tmpSql & "AND (ISNULL(APManifestNo, '') LIKE '%AP%' OR CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
        End If
    Else
        tmpSql = tmpSql & "And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
        If Check2.Value = 0 Then
            tmpSql = tmpSql & "AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
        End If
    End If
    
    tmpSql = tmpSql & "AND UnderSaleStaffID = '" & StaffID & "'" & vbCrLf
    If StaffType = 2 Then
        If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode <> "AKP" Then
            tmpSql = tmpSql & "AND SaleZoneID = '" & SaleZoneID & "'" & vbCrLf
        End If
    End If
    
    tmpSql = tmpSql & "AND SaleDistrictID IN (SELECT SaleDistrictID FROM tbSaleTargetDistrict S WHERE S.CompanyID + '_' + S.SaleDistrictID + '_' + CAST(S.ForYear AS Varchar(4)) + '_' + CAST(S.ForMonth AS Varchar(2)) = C.CompanyID + '_' + C.SaleDistrictID + '_' + CAST(C.ForYear AS Varchar(4)) + '_' + CAST(C.ForMonth AS Varchar(2)))"
    
'    tmpSql = tmpSql & "AND isJobs = '" & isJob & "'" & vbCrLf
    tmpSql = tmpSql & "GROUP BY CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth, DocNo"
    
    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & tmpTable_CommFee & "]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[" & tmpTable_CommFee & "]"
    DBConnExc tSql
    DBConnExc tmpSql
    DoEvents
    
    'Group ยอดขายใบเมนิเฟสตาม ปี - เดือน
    tmpSql = "SELECT CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForMonth, ForYear, SUM(InCollectFee) AS InCollectFee, SUM(TripTransFee) AS TripTransFee, SUM(TripTransFee_Inc) AS TripTransFee_Inc, SUM(VDTransFee) AS VDTransFee, SUM(TripTransFeeDiff) AS TripTransFeeDiff, SUM(NonFee) AS NonFee, " & vbCrLf
    tmpSql = tmpSql & "SUM(HazFee) AS HazFee, SUM(NBLFee) AS NBLFee, SUM(HBLFee) AS HBLFee, SUM(HICFee) AS HICFee, SUM(HBIFee) AS HBIFee, SUM(BBHFee) AS BBHFee, SUM(BBLFee) AS BBLFee, SUM(NonQty) AS NonQty, SUM(NBLQty) AS NBLQty, SUM(HazQty) AS HazQty, SUM(HBLQty) AS HBLQty, SUM(HICQty) AS HICQty, SUM(HBIQty) AS HBIQty, SUM(BBHQty) AS BBHQty, SUM(BBLQty) AS BBLQty" & vbCrLf
    tmpSql = tmpSql & "From " & tmpTable_CommFee & vbCrLf
    tmpSql = tmpSql & "GROUP BY CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth"
    
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tmpSql, adOpenForwardOnly, adLockReadOnly)
        
    DBConnExc "DELETE FROM " & tmpTable_CommFee
    ProgressBar1.Value = 0
    ProgressBar1.Min = 0
    If RS.RecordCount > 0 Then
        ProgressBar1.Max = RS.RecordCount
    Else
        ProgressBar1.Max = 1
    End If
    Do While Not RS.EOF
        ProgressBar1.Max = RS.RecordCount
        ProgressBar1.Value = RS.AbsolutePosition
        DoEvents
        tmpSql = "SELECT * FROM " & tmpTable_CommFee & vbCrLf
        tmpSql = tmpSql & "WHERE CompanyID = '" & Trim(RS!CompanyID) & "' AND UnderSaleStaffID = '" & Trim(RS!UnderSaleStaffID) & "'" & vbCrLf
        tmpSql = tmpSql & "AND ForYear = '" & Trim(RS!ForYear) & "' AND ForMonth = '" & Trim(RS!ForMonth) & "'"
        Set RSInsert = New ADODB.Recordset
        GetDBConnection
        RSInsert.Open tmpSql, DBConn, adOpenKeyset, adLockOptimistic
        If RSInsert.EOF Then
            RSInsert.AddNew
            RSInsert!CompanyID = RS!CompanyID
            RSInsert!SaleZoneID = RS!SaleZoneID
            RSInsert!SaleDistrictID = RS!SaleDistrictID
            RSInsert!UnderSaleStaffID = RS!UnderSaleStaffID
            RSInsert!ForMonth = RS!ForMonth
            RSInsert!ForYear = RS!ForYear
            RSInsert!InCollectFee = RS!InCollectFee
            RSInsert!TripTransFee = RS!TripTransFee
            RSInsert!TripTransFee_Inc = RS!TripTransFee_Inc
            RSInsert!VDTransFee = RS!VDTransFee
            RSInsert!TripTransFeeDiff = RS!TripTransFeeDiff
            RSInsert!NonFee = RS!NonFee
            RSInsert!HazFee = RS!HazFee
            RSInsert!NBLFee = RS!NBLFee
            RSInsert!HBLFee = RS!HBLFee
            RSInsert!HICFee = RS!HICFee
            RSInsert!HBIFee = RS!HBIFee
            RSInsert!BBHFee = RS!BBHFee
            RSInsert!BBLFee = RS!BBLFee
            RSInsert!NonQty = RS!NonQty
            RSInsert!NBLQty = RS!NBLQty
            RSInsert!HazQty = RS!HazQty
            RSInsert!HBLQty = RS!HBLQty
            RSInsert!HICQty = RS!HICQty
            RSInsert!HBIQty = RS!HBIQty
            RSInsert!BBHQty = RS!BBHQty
            RSInsert!BBLQty = RS!BBLQty
            RSInsert!ExtraRate = 0
            RSInsert.Update
        End If
        RS.MoveNext
    Loop
    CloseDBConnection
    
    tmpSql = "SELECT * FROM " & tmpTable_CommFee & " ORDER BY ForYear, ForMonth"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tmpSql, adOpenForwardOnly, adLockReadOnly)
    ProgressBar1.Value = 0
    ProgressBar1.Min = 0
    If RS.RecordCount > 0 Then
        ProgressBar1.Max = RS.RecordCount
    Else
        ProgressBar1.Max = 1
    End If
    'Update ข้อมูลยอดขายในเมนิเฟส ลงใน tbCom_Commission
    Do While Not RS.EOF
        ProgressBar1.Max = RS.RecordCount
        ProgressBar1.Value = RS.AbsolutePosition
        DoEvents
        Deduck = 0
        If StaffType = 1 Then
            For i = 1 To ctlGrid3.getGridObj.Rows - 1
                If ctlGrid3.getGridObj.TextMatrix(i, 0) = RS!ForMonth Then
                    Deduck = IIf(IsNumeric(ctlGrid3.getGridObj.TextMatrix(i, 1)), ctlGrid3.getGridObj.TextMatrix(i, 1), 0)
                    Exit For
                End If
            Next
        End If
        tmpSql = "UPDATE " & tmpTable_Comm & " SET InCollectFee = " & IIf(IsNull(RS!InCollectFee), 0, RS!InCollectFee) - Deduck & " , NonFee = " & IIf(IsNull(RS!NonFee), 0, RS!NonFee) & "," & vbCrLf
        tmpSql = tmpSql & "HazFee = " & IIf(IsNull(RS!HazFee), 0, RS!HazFee) & ", NBLFee = " & IIf(IsNull(RS!NBLFee), 0, RS!NBLFee) & ", HBLFee = " & IIf(IsNull(RS!HBLFee), 0, RS!HBLFee) & ", " & vbCrLf
        tmpSql = tmpSql & "HICFee = " & IIf(IsNull(RS!HICFee), 0, RS!HICFee) & ", HBIFee = " & IIf(IsNull(RS!HBIFee), 0, RS!HBIFee) & ", BBHFee = " & IIf(IsNull(RS!BBHFee), 0, RS!BBHFee) & ", BBLFee = " & IIf(IsNull(RS!BBLFee), 0, RS!BBLFee) & ", " & vbCrLf
        tmpSql = tmpSql & "NonQty = " & IIf(IsNull(RS!NonQty), 0, RS!NonQty) & ", NBLQty = " & IIf(IsNull(RS!NBLQty), 0, RS!NBLQty) & "," & vbCrLf
        tmpSql = tmpSql & "HazQty = " & IIf(IsNull(RS!HazQty), 0, RS!HazQty) & ", HBLQty = " & IIf(IsNull(RS!HBLQty), 0, RS!HBLQty) & "," & vbCrLf
        tmpSql = tmpSql & "HICQty = " & IIf(IsNull(RS!HICQty), 0, RS!HICQty) & ", HBIQty = " & IIf(IsNull(RS!HBIQty), 0, RS!HBIQty) & vbCrLf
        tmpSql = tmpSql & "WHERE SaleStaffID = '" & RS!UnderSaleStaffID & "'" & vbCrLf
        tmpSql = tmpSql & "AND SaleDistrictID = '" & RS!SaleDistrictID & "'" & vbCrLf
        
        If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode <> "AKP" Then
            tmpSql = tmpSql & "AND  SaleZoneID = '" & Trim(RS!SaleZoneID) & "'" & vbCrLf
        End If
'        tmpSql = tmpSql & "AND  CompanyID = '" & Trim(RS!CompanyID) & "'" & vbcrlf
        tmpSql = tmpSql & "AND ForYear = '" & RS!ForYear & "' AND ForMonth = " & RS!ForMonth
        
        DBConnExc tmpSql
        RS.MoveNext
    Loop
    Set RSInsert = Nothing
    Set RS = Nothing
    
    tSql = "SELECT * FROM " & tmpTable_Comm & " ORDER BY ForYear, ForMonth"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        ProgressBar1.Value = 0
        ProgressBar1.Min = 0
        If RS.RecordCount > 0 Then
            ProgressBar1.Max = RS.RecordCount
        Else
            ProgressBar1.Max = 1
        End If
        Do While Not RS.EOF
            ProgressBar1.Max = RS.RecordCount
            ProgressBar1.Value = RS.AbsolutePosition
            DoEvents
            If StaffType = 1 Then
                If ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 1) = Format(RS!ForYear, "0000") & Format(RS!ForMonth, "00") Then
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 3) = Format(RS!InCollectFee, "#,##0.00")
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 4) = Format((RS!NonFee + RS!NBLFee), "#,##0.00")
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 5) = Format((RS!HazFee + RS!HBLFee), "#,##0.00")
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 6) = Format((RS!NonQty + RS!NBLQty), "#,##0.00")
                    ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 7) = Format((RS!HazQty + RS!HBLQty), "#,##0.00")
                    If (RS!NonQty + RS!NBLQty) = 0 Then
                        ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 8) = "0.00"
                        ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 9) = "0.00"
                    Else
                        ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 8) = Format((RS!NonFee + RS!NBLFee) / (RS!NonQty + RS!NBLQty), "#,##0.00")
                        ctlGrid1.getGridObj.TextMatrix(RS.AbsolutePosition, 9) = Format((RS!HazFee + RS!HBLFee) / (RS!HazQty + RS!HBLQty), "#,##0.00")
                    End If
                    
                End If
            Else
                For i = 1 To ctlGrid1.getGridObj.Rows - 1
                    If ctlGrid1.getGridObj.TextMatrix(i, 1) = Format(RS!ForYear, "0000") & Format(RS!ForMonth, "00") Then
                        If IsNumeric(ctlGrid1.getGridObj.TextMatrix(i, 3)) = False Then
                            ManifestFee = 0
                        Else
                            ManifestFee = CDbl(ctlGrid1.getGridObj.TextMatrix(i, 3))
                        End If
                        ManifestFee = ManifestFee + RS!InCollectFee
                        ctlGrid1.getGridObj.TextMatrix(i, 3) = Format(ManifestFee, "#,##0.00")
                        Exit For
                    Else
                        ManifestFee = 0
                    End If
                Next
            End If
            RS.MoveNext
        Loop
    End If
    Set RSInsert = Nothing
    Set RS = Nothing
    'ค้นหารค่าขนส่งและหักออกจากยอดขายในใบเมนิเฟส
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
        Call CalTimeTrans(StaffType, StaffID, SaleZoneID, isJob, Check3.Value)
    Else
        Call CalTimeTrans(StaffType, StaffID, SaleZoneID, isJob, Check2.Value)
    End If
    Exit Function
ErrD:
    VerifyData = False
    Exit Function
End Function

Private Sub CalTimeTrans(ByVal StaffType As Integer, ByVal StaffID As String, Optional ByVal SaleZoneID As String = "", Optional ByVal isJob As String = "N", Optional ByVal isCover As Integer = 0)
Dim RS As ADODB.Recordset
Dim RSInsert As ADODB.Recordset
Dim tmpSql As String
Dim ManifestFee As Double, NonFee As Double, HazFee As Double, i As Integer
Dim TripTransFee As Double
Dim TempTable As String
    
    TempTable = "TEMPCommission_TransFee_" & ComName
    '*** Get Trans Fee for Include Trans
'    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
'        tmpSql = "EXEC sp_GetTransport_Fee_Inc '" & dtFrom.ValueYMD & "', '" & dtTo.ValueYMD & "', '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "', 'Y', '" & tmpTable_CommFee & "','" & isCover & "'"
'    Else
'        tmpSql = "EXEC sp_GetTransport_Fee_Inc '" & dtFrom.ValueYMD & "', '" & dtTo.ValueYMD & "', '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "', 'N', '" & tmpTable_CommFee & "','" & isCover & "'"
'    End If
'    DBConnExc tmpSql
       
    tmpSql = "SELECT CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForMonth, ForYear, MAX(TripTransFee) AS TripTransFee, MAX(TripTransFee_Inc) AS TripTransFee_Inc, MAX(VDTransFee) AS VDTransFEE," & vbCrLf
    tmpSql = tmpSql & "CASE" & vbCrLf
    tmpSql = tmpSql & "    WHEN MAX(TripTransFee) > 0 THEN" & vbCrLf
    tmpSql = tmpSql & "                                    CASE WHEN MAX(TripTransFee) < MAX(VDTransFee) THEN  MAX(VDTransFee) - MAX(TripTransFee) ELSE 0 END" & vbCrLf
    tmpSql = tmpSql & "    WHEN MAX(TripTransFee) = 0 THEN 0" & vbCrLf
    tmpSql = tmpSql & "End AS TripTransFeeDiff" & vbCrLf
    tmpSql = tmpSql & "INTO " & TempTable & vbCrLf
    tmpSql = tmpSql & "From dbo.Com_vwCommission_Fee" & vbCrLf
    tmpSql = tmpSql & "Where MinWorkDate >= '" & dtFrom.ValueYMD & "'"
    tmpSql = tmpSql & " AND MaxWorkDate <= '" & dtTo.ValueYMD & "'" & vbCrLf
    
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID = "776EDA5830F744A6B3551348D851BC22" Then
        If Check3.Value = 0 Then
            tmpSql = tmpSql & "AND (ISNULL(APManifestNo, '') LIKE '%AP%' OR CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
        End If
    Else
        tmpSql = tmpSql & "And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
        If Check2.Value = 0 Then
            tmpSql = tmpSql & "AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
        End If
    End If
    
    tmpSql = tmpSql & "AND UnderSaleStaffID = '" & StaffID & "'" & vbCrLf
    If StaffType = 2 Then
        If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID <> "776EDA5830F744A6B3551348D851BC22" Then
            tmpSql = tmpSql & "AND SaleZoneID = '" & SaleZoneID & "'" & vbCrLf
        End If
    End If
'    tmpSql = tmpSql & "AND isJobs = '" & isJob & "'" & vbCrLf
    tmpSql = tmpSql & "GROUP BY CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth, TimeTableNo"
            
    tSql = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & TempTable & "]') AND type in (N'U'))" & vbCrLf
    tSql = tSql & "DROP TABLE [dbo].[" & TempTable & "]"
    DBConnExc tSql
    DBConn.Execute tmpSql
    
'    tmpSql = "SELECT CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForMonth, ForYear, SUM(TripTransFee) AS TripTransFee, SUM(TripTransFee_Inc) AS TripTransFee_Inc, SUM(VDTransFee) AS VDTransFEE, SUM(TripTransFeeDiff) AS TripTransFeeDiff" & vbCrLf
'    tmpSql = tmpSql & "FROM " & TempTable & vbCrLf
'    tmpSql = tmpSql & "GROUP BY CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth"

    tmpSql = "SELECT CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForMonth, ForYear, SUM(TripTransFee) AS TripTransFee, SUM(TripTransFee_Inc) AS TripTransFee_Inc, SUM(VDTransFee) AS VDTransFEE, SUM(TripTransFeeDiff) AS TripTransFeeDiff" & vbCrLf
    tmpSql = tmpSql & ", CompanyID + '_' + SaleDistrictID + '_' + CAST(ForYear AS Varchar(4)) + '_' + CAST(ForMonth AS Varchar(2))" & vbCrLf
    tmpSql = tmpSql & "FROM " & TempTable & " C" & vbCrLf
    tmpSql = tmpSql & "WHERE SaleDistrictID IN (SELECT SaleDistrictID FROM tbSaleTargetDistrict S WHERE s.CompanyID + '_' + s.SaleDistrictID + '_' + CAST(s.ForYear AS Varchar(4)) + '_' + CAST(s.ForMonth AS Varchar(2)) = C.CompanyID + '_' + C.SaleDistrictID + '_' + CAST(C.ForYear AS Varchar(4)) + '_' + CAST(C.ForMonth AS Varchar(2)))" & vbCrLf
    tmpSql = tmpSql & "GROUP BY CompanyID, SaleZoneID, SaleDistrictID, UnderSaleStaffID, ForYear, ForMonth"

    Set RS = New ADODB.Recordset
    Set RS = GetRS(tmpSql, adOpenForwardOnly, adLockReadOnly)
    ProgressBar1.Value = 0
    ProgressBar1.Min = 0
    If RS.RecordCount > 0 Then
        ProgressBar1.Max = RS.RecordCount
    Else
        ProgressBar1.Max = 1
    End If
    TripTransFee = 0
    Do While Not RS.EOF
        ProgressBar1.Max = RS.RecordCount
        ProgressBar1.Value = RS.AbsolutePosition
        DoEvents
        TripTransFee = TripTransFee + IIf(IsNull(RS!TripTransFee_Inc), 0, RS!TripTransFee_Inc)
            tmpSql = "UPDATE " & tmpTable_Comm & " SET TripTransFee = " & IIf(IsNull(RS!TripTransFee), 0, RS!TripTransFee) & vbCrLf
            tmpSql = tmpSql & ", TripTransFee_Inc = " & IIf(IsNull(RS!TripTransFee_Inc), 0, RS!TripTransFee_Inc) & vbCrLf
            tmpSql = tmpSql & ", VDTransFee= " & IIf(IsNull(RS!VDTransFee), 0, RS!VDTransFee) & vbCrLf
            tmpSql = tmpSql & ", TripTransFeeDiff = " & IIf(IsNull(RS!TripTransFeeDiff), 0, RS!TripTransFeeDiff) & vbCrLf
            tmpSql = tmpSql & "WHERE SaleStaffID = '" & RS!UnderSaleStaffID & "'" & vbCrLf
            tmpSql = tmpSql & "AND ForYear = '" & RS!ForYear & "' AND ForMonth = '" & RS!ForMonth & "'"
        DBConnExc tmpSql
        RS.MoveNext
    Loop
    Set RSInsert = Nothing
    Set RS = Nothing
      
    'Update ยอดขายใบเมนิเฟส
    ManifestFee = 0
    NonFee = 0
    HazFee = 0
    tSql = "SELECT * FROM " & tmpTable_Comm & " ORDER BY ForYear, ForMonth"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        Do While Not RS.EOF
            For i = 1 To ctlGrid1.getGridObj.Rows - 1
                If ctlGrid1.getGridObj.TextMatrix(i, 1) = Format(RS!ForYear, "0000") & Format(RS!ForMonth, "00") Then
                    If IsNumeric(ctlGrid1.getGridObj.TextMatrix(i, 3)) = False Then
                        ManifestFee = 0
                        NonFee = 0
                        HazFee = 0
                    Else
                        ManifestFee = CDbl(ctlGrid1.getGridObj.TextMatrix(i, 3))
                        NonFee = CDbl(ctlGrid1.getGridObj.TextMatrix(i, 4))
                        HazFee = CDbl(ctlGrid1.getGridObj.TextMatrix(i, 5))
                    End If
                    ManifestFee = ManifestFee - RS!TripTransFee_Inc - RS!TripTransFeeDiff
                    NonFee = NonFee - ((RS!TripTransFee_Inc - RS!TripTransFeeDiff) / 2)
                    HazFee = HazFee - ((RS!TripTransFee_Inc - RS!TripTransFeeDiff) / 2)
'                    If RS!TripTransFee - RS!VDTransFee < 0 Then ManifestFee = ManifestFee - (RS!TripTransFee - RS!VDTransFee)
                    ctlGrid1.getGridObj.TextMatrix(i, 3) = Format(ManifestFee, "#,##0.00")
                    ctlGrid1.getGridObj.TextMatrix(i, 4) = Format(NonFee, "#,##0.00")
                    ctlGrid1.getGridObj.TextMatrix(i, 5) = Format(HazFee, "#,##0.00")
                    
                    If Int(ctlGrid1.getGridObj.TextMatrix(i, 6)) = 0 Then
                        ctlGrid1.getGridObj.TextMatrix(i, 8) = "0.00"
                    Else
                        ctlGrid1.getGridObj.TextMatrix(i, 8) = Format(NonFee / ctlGrid1.getGridObj.TextMatrix(i, 6), "#,##0.00")
                    End If
                    
                    If Int(ctlGrid1.getGridObj.TextMatrix(i, 7)) = 0 Then
                        ctlGrid1.getGridObj.TextMatrix(i, 9) = "0.00"
                    Else
                        ctlGrid1.getGridObj.TextMatrix(i, 9) = Format(HazFee / ctlGrid1.getGridObj.TextMatrix(i, 7), "#,##0.00")
                    End If
                    
                    Exit For
                Else
                    ManifestFee = 0
                End If
            Next
            RS.MoveNext
        Loop
    End If
    Set RSInsert = Nothing
    Set RS = Nothing
    
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode = "AKP" Then
        tmpSql = "UPDATE " & tmpTable_Comm & " SET CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'"
        DBConnExc tmpSql
    End If
End Sub


