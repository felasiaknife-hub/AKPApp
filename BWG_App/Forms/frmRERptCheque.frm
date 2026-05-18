VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmRERptCheque 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   60
      TabIndex        =   5
      Top             =   840
      Width           =   5955
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cboAccNo 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   4455
      End
      Begin VB.ComboBox cboBank 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
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
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "/นำฝาก"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขึ้นเงิน"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่บัญชี"
         Height          =   195
         Index           =   3
         Left            =   390
         TabIndex        =   11
         Top             =   1380
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ธนาคาร"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   10
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   7
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   60
      Top             =   2700
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
      Caption         =   "รายงานเช็คนำฝากธนาคาร"
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
      TabIndex        =   4
      Top             =   60
      Width           =   2250
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frmRERptCheque.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmRERptCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sessionID_ As String
Private Type Bank
    BankID As String
    BankCode As String
    BankName As String
End Type
Private Type BankAcc
    BankAccNo As String
    BankAccName As String
End Type
Dim BankArr() As Bank
Dim BankAccArr() As BankAcc

Private Sub cboAccNo_DropDown()
    GetAccNo
End Sub

Private Sub cboBank_Validate(Cancel As Boolean)
    GetAccNo
End Sub

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
    PopulateCompany cboCompany, cboIndex
    GetBank
    GetAccNo
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = "01" '
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, iCount As Long
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            tmpSql = "{vw_REReports.CompanyID} = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
        End If
    Else
        If cboCompany.ListIndex >= 0 Then
            tmpSql = "{vw_REReports.CompanyID} = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
        End If
    End If
    
    tmpSql = tmpSql & "AND {tbReCheqe.DueDate} >= '" & dtFrom.ValueYMD & "' AND {tbReCheqe.DueDate} <= '" & dtTo.ValueYMD & "' "
    If cboBank.ListIndex > 0 Then
        tmpSql = tmpSql & "AND {tbReCheqe.ComBank} = '" & BankArr(cboBank.ListIndex).BankCode & "' "
        If cboAccNo.ListIndex > 0 Then
            tmpSql = tmpSql & "AND {tbReCheqe.ComAccount} = '" & BankAccArr(cboAccNo.ListIndex).BankAccNo & "' "
        End If
    End If

    With rpt1
        .Reset
        .WindowTitle = "รายงานใบสำคัญรับ"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\RERptCheque.rpt"
        .Formulas(0) = "Company = 'บริษัท " & Trim(cboCompany.Text) & "'"
        .Formulas(1) = "REDate = '" & "วันที่  " & dtFrom.ValueDMY & "  ถึงวันที่  " & dtTo.ValueDMY & "'"
        .SelectionFormula = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBConnExc "Delete tbPrintOrd01 WHERE SessionID='" & sessionID_ & "'"
End Sub



Private Sub GetBank()
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "SELECT * FROM tbBank ORDER BY BankCode"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenKeyset, adLockOptimistic)
    ReDim BankArr(RS.RecordCount)
    cboBank.Clear
    cboBank.AddItem "== เลือกทุกธนาคาร =="
    Do While Not RS.EOF
        BankArr(RS.AbsolutePosition).BankID = RS!BankID
        BankArr(RS.AbsolutePosition).BankCode = RS!BankCode
        BankArr(RS.AbsolutePosition).BankName = RS!BankName
        cboBank.AddItem RS!BankName
        RS.MoveNext
    Loop
    Set RS = Nothing
    If cboBank.ListCount > 0 Then cboBank.ListIndex = 0
End Sub

Private Sub GetAccNo()
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "SELECT * FROM tbBankAccount"
    If cboBank.ListIndex > 0 Then
        tSql = tSql & " WHERE BankID = '" & BankArr(cboBank.ListIndex).BankID & "'"
    End If
    tSql = tSql & " ORDER BY AccountNo"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenKeyset, adLockOptimistic)
    ReDim BankAccArr(RS.RecordCount)
    cboAccNo.Clear
    cboAccNo.AddItem "== เลือกทุกเลขที่บัญชี =="
    cboAccNo.ListIndex = 0
    If cboBank.ListIndex <= 0 Then Exit Sub
    Do While Not RS.EOF
        BankAccArr(RS.AbsolutePosition).BankAccNo = RS!AccountNo
        BankAccArr(RS.AbsolutePosition).BankAccName = RS!AccountName
        cboAccNo.AddItem RS!AccountNo
        RS.MoveNext
    Loop
    Set RS = Nothing
    If cboAccNo.ListCount > 0 Then cboAccNo.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage True
End Sub
