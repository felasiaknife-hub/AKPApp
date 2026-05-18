VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmComInt2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   5955
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   495
         TabIndex        =   32
         Top             =   2745
         Width           =   5145
         Begin VB.OptionButton Option1 
            Caption         =   "งาน Routine"
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   180
            Width           =   1365
         End
         Begin VB.OptionButton Option2 
            Caption         =   "งาน Job"
            Height          =   195
            Left            =   1620
            TabIndex        =   35
            Top             =   180
            Width           =   1185
         End
         Begin VB.OptionButton Option3 
            Caption         =   "งาน DOC"
            Height          =   195
            Left            =   2835
            TabIndex        =   34
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "ทั้งหมด"
            Height          =   195
            Left            =   4005
            TabIndex        =   33
            Top             =   180
            Width           =   960
         End
      End
      Begin VB.ComboBox CboFD 
         Height          =   315
         ItemData        =   "frmComInt2.frx":0000
         Left            =   1095
         List            =   "frmComInt2.frx":0013
         TabIndex        =   30
         Text            =   "=Non="
         Top             =   4365
         Width           =   975
      End
      Begin VB.ComboBox Cbo90day 
         Height          =   315
         ItemData        =   "frmComInt2.frx":0031
         Left            =   1080
         List            =   "frmComInt2.frx":003E
         TabIndex        =   29
         Text            =   "=Non="
         Top             =   3990
         Width           =   3255
      End
      Begin VB.ComboBox CboDF 
         Height          =   315
         ItemData        =   "frmComInt2.frx":007F
         Left            =   1080
         List            =   "frmComInt2.frx":0098
         TabIndex        =   26
         Text            =   "=Non="
         Top             =   3630
         Width           =   975
      End
      Begin VB.OptionButton OpOld 
         BackColor       =   &H80000005&
         Caption         =   "เขตการขายเก่า"
         Height          =   315
         Left            =   2880
         TabIndex        =   24
         Top             =   2925
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OpNow 
         BackColor       =   &H80000005&
         Caption         =   "เขตการขายปัจจุบัน"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   2940
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Text            =   "*"
         Top             =   2400
         Width           =   4455
      End
      Begin VB.ComboBox cboWaste 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   4455
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Text            =   "*"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   4455
      End
      Begin Commission.ctlDate dtFrom 
         Height          =   375
         Left            =   1140
         TabIndex        =   1
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate dtTo 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate ctlDate1 
         Height          =   375
         Left            =   1140
         TabIndex        =   3
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate ctlDate2 
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.Frame Frame3 
         Height          =   465
         Left            =   495
         TabIndex        =   37
         Top             =   3150
         Width           =   5145
         Begin VB.OptionButton Option8 
            Caption         =   "ทั้งหมด"
            Height          =   195
            Left            =   4005
            TabIndex        =   40
            Top             =   180
            Width           =   960
         End
         Begin VB.OptionButton Option6 
            Caption         =   "งานลูกค้าเก่า"
            Height          =   195
            Left            =   2025
            TabIndex        =   39
            Top             =   180
            Width           =   1545
         End
         Begin VB.OptionButton Option5 
            Caption         =   "งานลูกค้าใหม่"
            Height          =   195
            Left            =   180
            TabIndex        =   38
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วัน นับจากวันที่วางบิล ถึง วันที่ชำระ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2115
         TabIndex        =   41
         Top             =   4410
         Width           =   2430
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "นับวันเกิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   31
         Top             =   4365
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกเฉพาะ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   28
         Top             =   4020
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วัน นับจากวันที่วางบิล ถึง วันที่ชำระ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2115
         TabIndex        =   27
         Top             =   3630
         Width           =   2430
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "นับวันไม่เกิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   3630
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   22
         Top             =   1020
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน"
         Height          =   195
         Index           =   0
         Left            =   570
         TabIndex        =   21
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   20
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   6
         Left            =   3585
         TabIndex        =   19
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่รับชำระ"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waste"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   17
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   16
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   780
         TabIndex        =   15
         Top             =   1380
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   1740
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5775
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5775
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5775
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   60
      Top             =   5220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   5910
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานการรับชำระค่าบริการ / การตัดจ่ายค่าคอมฯ ภายใน"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   150
      TabIndex        =   12
      Top             =   60
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmComInt2.frx":00BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmComInt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_ As String
Dim curCustomer_ As BWGCustomer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%
    sessionID_ = GetGUID
    PopulateCompany
    PopulateZone
    InitWaste
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = "01"
    ctlDate1.DayValue = "01"
    Option4.Value = True
    Option8.Value = True
End Sub

Private Sub PrintMain(printType As Integer)
Dim RS As New ADODB.Recordset
Dim tmpSql As String, iCount As Long
Dim tSql As String, ExcView As String, ExcTable As String

Dim TableName As String
Dim Ds As String
    TableName = "Com_tbComIn22_" & ComName

    tmpSql = ""
    tmpSql = "SELECT * INTO " & TableName & " FROM Com_vwComIn22 " & vbCrLf
    tmpSql = tmpSql & "WHERE WorkDate >= '" & ctlDate1.ValueYMD & "' AND WorkDate <= '" & ctlDate2.ValueYMD & "'" & vbCrLf
    If CboDF.Text <> "=Non=" Then
    tmpSql = tmpSql & " and DFDate2 <= '" & CboDF.Text & "' " & vbCrLf
    End If
    If CboFD.Text <> "=Non=" Then
    tmpSql = tmpSql & " and DFDate2 > '" & CboFD.Text & "' " & vbCrLf
    End If
    If Option2.Value = True Then
        tmpSql = tmpSql & " And jobisjobs='Y'" & vbCrLf
    End If
    If Option1.Value = True Then
        tmpSql = tmpSql & " And jobisjobs='N'" & vbCrLf
    End If
    If Option3.Value = True Then
        tmpSql = tmpSql & " And jobisjobs='D'" & vbCrLf
    End If
    If Option4.Value = True Then
        tmpSql = tmpSql & " And (jobisjobs='D' or jobisjobs='N' or jobisjobs='Y')" & vbCrLf
    End If
    If Option5.Value = True Then
        tmpSql = tmpSql & " And (newcustjob='Y')" & vbCrLf
    End If
    If Option6.Value = True Then
        tmpSql = tmpSql & " And (newcustjob<>'Y')" & vbCrLf
    End If
    If Option8.Value = True Then
    End If

    If Cbo90day.Text = "ลูกค้าใหม่กำหนด 90 วัน" Then
         tmpSql = tmpSql & " and  SetDay = '90' " & vbCrLf
    ElseIf Cbo90day = "ไม่ใช่ลูกค้าใหม่กำหนด 90 วัน" Then
        tmpSql = tmpSql & " and  SetDay <> '90' " & vbCrLf
    End If
    
'    If Chkis90.Value = 1 Then
'     tmpSql = tmpSql & " and  SetDay = '90' " & vbCrLf
'    End If
'    tmpSql = tmpSql & "AND DueDate >= '" & dtFrom.ValueYMD & "' AND DueDate <='" & dtTo.ValueYMD & "'" & vbCrLf

    If cboZone.ListIndex > 0 Then
        tmpSql = tmpSql & "AND SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "'"
    End If
    If Trim(txtSale.Text) <> "" Then
        If Trim(txtSale.Text) <> "*" Then
           'If ctlDate1.ValueYMD < "2555/03/01" Then
            tmpSql = tmpSql & "AND UnderSaleStaffID='" & curStaff_.ID & "'" & vbCrLf
           'Else
            'tmpSql = tmpSql & "AND UnderSaleStaffID='" & curStaff_.ID & "' And isActive = 'Y'" & vbCrLf
           'End If
        End If
        If OpNow.Value = True Then
            tmpSql = tmpSql & " AND isActive = 'Y'" & vbCrLf
        Else
            tmpSql = tmpSql & " AND isActive <> 'Y'" & vbCrLf
        End If
    End If
    If cboWaste.ListIndex > 0 Then
        tmpSql = tmpSql & " AND NonOrHz='" & Trim(cboWaste.Text) & "'" & vbCrLf
    End If
    If Trim(txtCustomer.Text) <> "" Then
        If Trim(txtCustomer.Text) <> "*" Then
            tmpSql = tmpSql & " AND CustomerID = '" & txtCustomer.Tag & "'" & vbCrLf
        End If
    End If
       
    'ตรวจสอบว่า พนักงานขาย มีเป้าของ บริษัทนั้นๆ หรือ ไม่
    tSql = "SELECT * FROM Com_vwSales_District_Target" & vbCrLf
    tSql = tSql & "WHERE (CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
    If Not curStaff_ Is Nothing Then tSql = tSql & "AND SaleStaffID = '" & curStaff_.ID & "'" & vbCrLf
    tSql = tSql & "AND ForYear = " & dtFrom.YearValue & " AND ForMonth = " & dtFrom.MonthValue & vbCrLf
    tSql = tSql & "ORDER BY SDName"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode = "AKP" Then
        tmpSql = tmpSql & "AND (CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
        If Not RS.EOF Then
            If RS!IsOverall = "0" Then tmpSql = tmpSql & "AND NonOrHz IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG')" & vbCrLf
        Else
            tmpSql = tmpSql & "AND NonOrHz IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG')" & vbCrLf
        End If
    Else
        tmpSql = tmpSql & "AND (CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
        If Not RS.EOF Then
            If RS!IsOverall = "0" Then
                tmpSql = tmpSql & "AND NonOrHz NOT IN ('HBI', 'NIC', 'NBI', 'BBH','DOC','BWG')" & vbCrLf
            Else
                tmpSql = tmpSql & "AND NonOrHz IN ('HBI', 'NIC', 'NBI', 'BBH','DOC','BWG')" & vbCrLf
            End If
        Else
            tmpSql = tmpSql & "AND NonOrHz NOT IN ('HBI', 'NIC', 'NBI', 'BBH','DOC','BWG')" & vbCrLf
        End If
    End If
    
    '***Truck Type
'    tmpSql = tmpSql & "AND TruckSubTypeID IN (SELECT SubTypeID FROM tbTruckSubType WHERE TruckTypeID IN ( '67D1CCC9C8FF4782B89E3F56100A1592','D8B22F81A3D64E328407B736A27AEB3F' ))"


    ExcTable = "IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & TableName & "]'))" & vbCrLf
    ExcTable = ExcTable & "DROP TABLE [dbo].[" & TableName & "]"
    
    DBConn.Execute ExcTable
    
'    ExcView = "CREATE TABLE [dbo].[" & ViewName & "] AS " & vbCrLf
'    ExcView = ExcView & tmpSql
'    DBConn.Execute ExcView
    
    DBConn.Execute tmpSql
    
    tmpSql = "SELECT * " & vbCrLf
    tmpSql = tmpSql & "From " & TableName & " AS Com_vwComIn22" & vbCrLf
    tmpSql = tmpSql & "WHERE DueDate >= '" & dtFrom.ValueYMD & "' AND DueDate <='" & dtTo.ValueYMD & "'" & vbCrLf
    tmpSql = tmpSql & "ORDER BY WorkDate, TimeTableNo, ZoneName,StaffFName,StaffLName, CompanyName, WasteType,InCollectFee"
    
    With rpt1
        .Reset
        .WindowTitle = "รายงานการรับชำระค่าบริการ / การตัดจ่ายค่าคอมฯ ภายใน"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If dtTo.YearValue >= "2564" Then
        .ReportFileName = App.Path & "\Reports\ComIn2.rpt"
        Else
        .ReportFileName = App.Path & "\Reports\ComIn2_Old.rpt"
        End If
        .Formulas(0) = "Company='บริษัท " & Trim(cboCompany.Text) & "'"
        .Formulas(1) = "WorkDate='" & "วันที่รับชำระ  " & dtFrom.ValueDMY & "  ถึงวันที่  " & dtTo.ValueDMY & "'"
        .Formulas(2) = "Sales='" & IIf(Trim(txtSale.Text) <> "", IIf(Trim(txtSale) <> "*", txtSale.Text, ""), "") & "'"
        .SqlQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
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

Private Sub PopulateZone()
Dim x%, tmpS As New BWGSearchManager
    cboZone.Clear
    cboZone.AddItem "ทุกโซน"
    Set colZone_ = tmpS.SaleZoneSearch
    For x = 1 To colZone_.Count
        cboZone.AddItem colZone_(x).ZoneName
    Next
    If cboZone.ListCount > 0 Then
        cboZone.ListIndex = 0
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBConnExc "Delete tbPrintOrd01 Where SessionID='" & sessionID_ & "'"
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If txtCustomer.Text = "" Then Exit Sub
    If txtCustomer.Text = "*" Then Exit Sub
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    txtCustomer.Tag = ""
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frCustomerSearch
            With frCustomerSearch
                .getCriteria (txtCustomer.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
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

Private Sub InitWaste()
Dim x%
Dim strWaste As String
Dim tWaste() As String
    strWaste = "ไม่ระบุ,Haz,Non,HBL,NBL"
    tWaste = Split(strWaste, ",")
    cboWaste.Clear
    For x = 0 To 4
        cboWaste.AddItem tWaste(x)
    Next
    If cboWaste.ListCount > 0 Then
        cboWaste.ListIndex = 0
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    txtCustomer.Tag = ""
    With curCustomer_
        txtCustomer.Text = .CustomerName
        txtCustomer.Tag = .ID
    End With
End Sub

