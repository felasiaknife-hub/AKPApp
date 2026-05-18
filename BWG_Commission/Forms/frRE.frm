VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form frmRE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9840
      TabIndex        =   40
      Top             =   8040
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เช็ค"
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   1
      Left            =   60
      TabIndex        =   36
      Top             =   6180
      Width           =   7935
      Begin Project1.ctlGrid grCheqe 
         Height          =   1875
         Left            =   90
         TabIndex        =   37
         Top             =   240
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   3307
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ลบ"
      Height          =   795
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8520
      Width           =   1035
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9840
      TabIndex        =   32
      Top             =   7260
      Width           =   1515
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9840
      TabIndex        =   30
      Top             =   6900
      Width           =   1515
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   7860
      TabIndex        =   28
      Top             =   2460
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   1440
      TabIndex        =   27
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   1440
      TabIndex        =   26
      Top             =   2040
      Width           =   10275
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   1440
      TabIndex        =   25
      Top             =   1680
      Width           =   7695
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   1440
      TabIndex        =   22
      Top             =   1320
      Width           =   3855
   End
   Begin Project1.ctlDate dtBillDate 
      Height          =   375
      Left            =   4020
      TabIndex        =   21
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9840
      TabIndex        =   19
      Top             =   6540
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "เลือกใบวางบิล"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8520
      Width           =   1455
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   8220
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   1035
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใบใหม่"
      Height          =   795
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ใบวางบิลที่เลือก"
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   3000
      Width           =   11715
      Begin Project1.ctlGrid grItem 
         Height          =   2775
         Left            =   90
         TabIndex        =   20
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4895
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   4275
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   960
      Width           =   1875
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   2100
      Top             =   8700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผลต่างจากการเก็บเงิน"
      Height          =   195
      Index           =   14
      Left            =   8100
      TabIndex        =   42
      Top             =   8100
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   13
      Left            =   11460
      TabIndex        =   41
      Top             =   8100
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมยอดเงิน"
      Height          =   195
      Index           =   6
      Left            =   8820
      TabIndex        =   39
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAX."
      Height          =   195
      Index           =   9
      Left            =   8820
      TabIndex        =   38
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   12
      Left            =   11460
      TabIndex        =   34
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   11
      Left            =   8760
      TabIndex        =   33
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   10
      Left            =   11460
      TabIndex        =   31
      Top             =   6960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   7
      Left            =   11460
      TabIndex        =   29
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสำคัญรับ"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   1020
      Width           =   1275
   End
   Begin VB.Label lbCancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9780
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกเครื่องพิมพ์"
      Height          =   195
      Index           =   8
      Left            =   6960
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   8460
      Y2              =   8460
   End
   Begin VB.Label lbNetPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9900
      TabIndex        =   15
      Top             =   7740
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   5
      Left            =   11460
      TabIndex        =   14
      Top             =   7740
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมทั้งสิ้น"
      Height          =   195
      Index           =   4
      Left            =   8640
      TabIndex        =   13
      Top             =   7740
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบสำคัญรับ"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   120
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   180
      X2              =   6990
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รับในนามบริษัท"
      Height          =   195
      Index           =   1
      Left            =   5880
      TabIndex        =   10
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   2
      Left            =   3540
      TabIndex        =   9
      Top             =   1020
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frRE.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   3
      Left            =   9840
      Top             =   7680
      Width           =   1515
   End
End
Attribute VB_Name = "frmRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim myCR03Header_ As BWGPrintCR03
Dim curSignStaff_ As BWGCompanyStaff
Dim isDetailChange_ As Boolean

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex > 0 Then
        myCR03Header_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID
        myCR03Header_.CompanyName = cboCompany.Text
    Else
        myCR03Header_.CompanyID = ""
        myCR03Header_.CompanyName = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ท่านต้องการยกเลิกใบแจ้งการกำจัดกากนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myCR03Header_.CancelBill
        SetCancelScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCR03Header_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewCR03
    dtBillDate.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCR03Header_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frPrintCR03Search
    With frPrintCR03Search
        .setCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdPrint_Click()
Dim tmpSql As String, tmpRec As New ADODB.Recordset
Dim tmpCon As ADODB.Connection, iCount%, x%, tmpCol As New Collection
Dim selResult As Integer
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCR03Header_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        Set tmpCon = DBManager.GetDBConnection
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\reports\CR03.rpt"
        tmpSql = ""
'        tmpSql = "Select * from vw_CR03DetailSearch " & vbCrLf
'        tmpSql = tmpSql & "Where CR03ID='" & myCR03Header_.ID & "' "
'        tmpSql = tmpSql & vbCrLf & " Order By RowPos"
        tmpSql = "{vw_CR03DetailSearch.CR03ID} = '" & myCR03Header_.ID & "' "
        
        .Formulas(0) = "ShowCompanyName=" & Chr(34) & "ช่วงวันที่ " & myCR03Header_.CompanyName & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & "วันที่ " & dtBillDate.DayValue & "  เดือน " & dtBillDate.MonthNameShort & "  พ.ศ. " & dtBillDate.YearValue & Chr(34)
        .Formulas(2) = "ShowDocNo=" & Chr(34) & txtBillNo.Text & Chr(34)
        .Formulas(4) = "ShowSignName=" & Chr(34) & txtStaff.Text & Chr(34)
        If Not curSignStaff_ Is Nothing Then
            .Formulas(3) = "ShowSignPosition=" & Chr(34) & curSignStaff_.PositionName & Chr(34)
        End If
'        .SQLQuery = tmpSql
        .SelectionFormula = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdSave_Click()
    If cboCompany.ListIndex = -1 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบแจ้ง", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If myCR03Header_.PrintDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายละเอียดในใบแจ้ง", vbExclamation
        Exit Sub
    End If
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        myCR03Header_.Save
        txtBillNo.Text = myCR03Header_.CR03No
    End If
    If myCR03Header_.isChanged = False And myCR03Header_.isCancel = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    End If
End Sub

Private Sub Command2_Click()
    If Trim(cboCompany.Text) = "" Then
        MsgBox "กรุณาเลือกบริษัทก่อนทำการเลือกใบ Menifest", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmREIVAdd
    With frmREIVAdd
        .setCallerForm Me, CurrentUser.WorkCompanies(cboCompany.ListIndex + 1)
        .Show 1
        If isDetailChange_ Then
            ShowPrintDetail
        End If
    End With
End Sub

Private Sub dtBilldate_Validate(Cancel As Boolean)
    myCR03Header_.IssuedDate = dtBillDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%, selX As Integer
    PopulateCompany
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
        cmdDelete.Picture = .picDel
    End With
    With grItem.getGridObj
        .Cols = 4
        .TextMatrix(0, 1) = "ใบวางบิลที่เลือก"
        .TextMatrix(0, 2) = "คำอธิบาย"
        .TextMatrix(0, 3) = "จำนวนเงินตามใบวางบิล"
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 7150
        .ColWidth(3) = 2100
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
    End With
    With grItem
        .ColEnabled(1) = False
        .ColEnabled(2) = False
        .ColEnabled(3) = False
    End With
    
    With grCheqe.getGridObj
        .Cols = 5
        .TextMatrix(0, 1) = "เลขที่เช็ค"
        .TextMatrix(0, 2) = "ธนาคาร"
        .TextMatrix(0, 3) = "จำนวน"
        .TextMatrix(0, 4) = "วันที่ขึ้นเงิน"
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 1800
        .ColWidth(3) = 2000
        .ColWidth(4) = 1700
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
    End With
    With grCheqe
        .ColEnabled(1) = False
        .ColEnabled(2) = False
        .ColEnabled(3) = True
        .ColEnabled(4) = True
    End With
    CreateNewCR03
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myCR03Header_ Is Nothing Then Exit Sub
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myCR03Header_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
    grItem.ColEnabled(col) = (col = 7 And Trim(grItem.getGridObj.TextMatrix(row, 0)) <> "")
End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGPrintCR03Det
Dim tmpKey As String
    If col = 7 Then
        tmpKey = "'" & grItem.getGridObj.TextMatrix(row, 0) & "'"
        If isExist(myCR03Header_.PrintDetails, tmpKey) Then
            Set tmpDet = myCR03Header_.PrintDetails(tmpKey)
            tmpDet.CRRemDesc = Trim(grItem.getGridObj.Text)
            Set tmpDet = Nothing
            myCR03Header_.isChanged = True
        End If
    End If
End Sub

Private Sub txtStaff_Validate(Cancel As Boolean)
Dim selStr As String
    If curSignStaff_ Is Nothing Then Set curSignStaff_ = New BWGCompanyStaff
    selStr = Trim(curSignStaff_.StaffFName & " " & curSignStaff_.StaffLName)
    If Trim(txtStaff.Text) = "" And Trim(selStr) <> Trim(txtStaff.Text) Then
        Set curSignStaff_ = Nothing
        Set curSignStaff_ = New BWGCompanyStaff
        With myCR03Header_
            .SignStaffID = ""
            .SignStaffName = ""
        End With
    Else
        If Trim(selStr) <> Trim(txtStaff.Text) Then
            Set curSignStaff_ = Nothing
            Set curSignStaff_ = New BWGCompanyStaff
            With myCR03Header_
                .SignStaffID = ""
                .SignStaffName = ""
            End With
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtStaff.Text)
                .setCallerForm Me
                .Show 1
            End With
            If myCR03Header_.SignStaffID = "" Then txtStaff.Text = ""
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

Private Sub ShowPrintDetail()
Dim tmpDet As BWGPrintCR03Det
Dim x%
    grItem.ClearAllData
    With grItem.getGridObj
        .Rows = myCR03Header_.PrintDetails.Count + 20
        For x = 1 To myCR03Header_.PrintDetails.Count
            Set tmpDet = myCR03Header_.PrintDetails(x)
            .TextMatrix(x, 0) = tmpDet.ID
            .TextMatrix(x, 1) = FormatYMD_to_DMY(tmpDet.WorkDate)
            .TextMatrix(x, 2) = tmpDet.docNo
            .TextMatrix(x, 3) = tmpDet.WasteName
            .TextMatrix(x, 4) = tmpDet.WasteCode
            .TextMatrix(x, 5) = tmpDet.CarRegisID
            .TextMatrix(x, 6) = tmpDet.NetWeight
            .TextMatrix(x, 7) = tmpDet.CRRemDesc
            tmpDet.RowPos = x
            Set tmpDet = Nothing
        Next
    End With
    SumTotalWeight
End Sub

Private Sub CreateNewCR03()
    Set myCR03Header_ = New BWGPrintCR03
    With myCR03Header_
        .ID = GetGUID
        .IssuedDate = dtBillDate.ValueYMD
        If cboCompany.ListIndex > -1 Then
            .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID
            .CompanyName = cboCompany.Text
        End If
        .IssuedStaffID = CurrentUser.ID
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .isChanged = False
        SetCancelScreen
    End With
End Sub

Public Sub RemovePrintDetail(jobDetID As String, menifestID As String)
Dim tmpCol As Collection
Dim x%, tmpCol2 As Collection, tmpS As New BWGSearchManager
Dim tmpKey As String
    Set tmpCol = myCR03Header_.PrintDetails
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If Trim(tmpCol(x).menifestID) = Trim(menifestID) And Trim(tmpCol(x).jobDetID) = Trim(jobDetID) Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myCR03Header_.isChanged = True
    End If
    Set tmpS = Nothing
End Sub

Public Sub AddPrintDetail(selItem As BWGJobDataDetail)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String, TreatmentDesc As String
Dim tmpDetail As New BWGPrintCR03Det, tmpS As New BWGSearchManager
    Set tmpCol = tmpS.WasteTreatmentSearch("WasteDataID='" & selItem.WasteDataID & "'")
    If tmpCol.Count > 0 Then
        TreatmentDesc = tmpCol(1).TreatmentDesc
    Else
        TreatmentDesc = ""
    End If
    Set tmpCol = myCR03Header_.PrintDetails
    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        If tmpCol(x).menifestID = tmpID Then
            Set tmpDetail = tmpCol(x)
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpDetail.CR03ID = myCR03Header_.ID
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
    End If
    With tmpDetail
        .CarRegisID = selItem.MenifestCarRegisID
        .docNo = selItem.MenifestDocNo
        .jobDetID = selItem.ID
        .menifestID = selItem.menifestID
        If selItem.ChargeWeightType = "A" Then
            .NetWeight = selItem.DisposerWeight
        ElseIf selItem.ChargeWeightType = "B" Then
            .NetWeight = selItem.EstWasteQty
        End If
        .RowPos = tmpCol.Count
        .TreatmentDesc = TreatmentDesc
        .WasteCode = selItem.WasteEUCode
        .WasteName = selItem.WasteName
        .WorkDate = selItem.MenifestWorkDate
    End With
    isDetailChange_ = True
    myCR03Header_.isChanged = True
    Set tmpDetail = Nothing
    
    If myCR03Header_.PrintDetails Is Nothing Then
        Set myCR03Header_.PrintDetails = tmpCol
        isDetailChange_ = True
        myCR03Header_.isChanged = True
    End If
    Set tmpS = Nothing
    Set tmpCol = Nothing
End Sub

Public Function GetJobDetails() As Collection
Dim x%, tmpS As New BWGSearchManager
Dim tmpCol As New Collection
Dim tmpDet As BWGPrintCR03Det, tmpJobDet As BWGJobDataDetail
Dim selCol As Collection, tmpStr As String

    Set selCol = myCR03Header_.PrintDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            Set tmpJobDet = New BWGJobDataDetail
            tmpJobDet.ID = tmpDet.jobDetID
            tmpJobDet.menifestID = tmpDet.menifestID
            tmpJobDet.MenifestWorkDate = tmpDet.WorkDate
            tmpJobDet.MenifestDocNo = tmpDet.docNo
            tmpJobDet.WasteName = tmpDet.WasteName
            tmpJobDet.WasteEUCode = tmpDet.WasteCode
            tmpJobDet.MenifestCarRegisID = tmpDet.CarRegisID
            tmpJobDet.DisposerWeight = tmpDet.NetWeight
            tmpJobDet.EstWasteQty = tmpDet.NetWeight
            tmpCol.Add tmpJobDet, "'" & tmpJobDet.ID & "_" & tmpJobDet.menifestID & "'"
            Set tmpDet = Nothing
            Set tmpJobDet = Nothing
        Next
    End With
    Set GetJobDetails = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub ClearScreen()
    txtBillNo.Text = "== AUTO =="
    dtBillDate.ValueYMD = TodayDate
    grItem.ClearAllData
    grItem.getGridObj.Rows = 50
    lbNetPrice.Caption = "0.00"
    txtStaff.Text = ""
    Set curSignStaff_ = Nothing
    Set curSignStaff_ = New BWGCompanyStaff
End Sub

Private Sub SetCancelScreen()
Dim cVal As Boolean
    cVal = myCR03Header_.isCancel
    lbCancel.Visible = cVal
    cmdSave.Enabled = Not cVal
    Command2.Enabled = Not cVal
    cmdCancel.Enabled = Not cVal
End Sub

Public Sub SetDataCR03(selBill As BWGPrintCR03)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim x%
    Set myCR03Header_ = selBill
    With myCR03Header_
        txtBillNo.Text = .CR03No
        dtBillDate.ValueYMD = .IssuedDate
        SetTextToCombo .CompanyName, cboCompany
        Set tmpCol = tmpS.UserSearch("StaffID='" & .SignStaffID & "'")
        If tmpCol.Count > 0 Then
            SetDataStaff tmpCol(1), False
        End If
        ShowPrintDetail
    End With
    SetCancelScreen
End Sub

Public Property Get RefID() As String
    RefID = myCR03Header_.ID
End Property

Private Sub SumTotalWeight()
Dim x%, tmpAmt As Double
    With grItem.getGridObj
        For x = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(x, 6)) Then
                tmpAmt = tmpAmt + CDbl(.TextMatrix(x, 6))
            End If
        Next
        lbNetPrice.Caption = Format(tmpAmt, "#,##0.000")
    End With
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
    Set curSignStaff_ = selStaff
    With curSignStaff_
        txtStaff.Text = .StaffFName & " " & .StaffLName
        If isAssigned Then
            myCR03Header_.SignStaffName = .StaffFName & " " & .StaffLName
            myCR03Header_.SignStaffID = .ID
        End If
    End With
End Sub
