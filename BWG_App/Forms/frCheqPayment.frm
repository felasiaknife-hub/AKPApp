VERSION 5.00
Begin VB.Form frmCheqPayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "บันทึกการชำระเงิน"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WasteManagment.ctlDate dtPay 
      Height          =   315
      Left            =   1110
      TabIndex        =   38
      Top             =   150
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   795
      Left            =   5940
      Picture         =   "frCheqPayment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3960
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7050
      Picture         =   "frCheqPayment.frx":00E9
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3960
      Width           =   1065
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   7
      Left            =   6360
      TabIndex        =   35
      Top             =   3420
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   7
      Left            =   2430
      TabIndex        =   33
      Top             =   3420
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   7
      Left            =   330
      TabIndex        =   32
      Top             =   3420
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   6360
      TabIndex        =   31
      Top             =   3060
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   2430
      TabIndex        =   29
      Top             =   3060
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   330
      TabIndex        =   28
      Top             =   3060
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   6360
      TabIndex        =   27
      Top             =   2700
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   2430
      TabIndex        =   25
      Top             =   2700
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   330
      TabIndex        =   24
      Top             =   2700
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   6360
      TabIndex        =   23
      Top             =   2340
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   2430
      TabIndex        =   21
      Top             =   2340
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   330
      TabIndex        =   20
      Top             =   2340
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   6360
      TabIndex        =   19
      Top             =   1980
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2430
      TabIndex        =   17
      Top             =   1980
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   330
      TabIndex        =   16
      Top             =   1980
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   6360
      TabIndex        =   15
      Top             =   1620
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2430
      TabIndex        =   13
      Top             =   1620
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   330
      TabIndex        =   12
      Top             =   1620
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   6360
      TabIndex        =   11
      Top             =   1260
      Width           =   1545
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2430
      TabIndex        =   9
      Top             =   1260
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   330
      TabIndex        =   8
      Top             =   1260
      Width           =   2085
   End
   Begin VB.TextBox txtCheqAmt 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   6360
      TabIndex        =   7
      Top             =   900
      Width           =   1545
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   0
      Left            =   4560
      TabIndex        =   5
      Top             =   900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2430
      TabIndex        =   3
      Top             =   900
      Width           =   2085
   End
   Begin VB.TextBox txtCheqNo 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   900
      Width           =   2085
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   1
      Left            =   4560
      TabIndex        =   10
      Top             =   1260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   2
      Left            =   4560
      TabIndex        =   14
      Top             =   1620
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   3
      Left            =   4560
      TabIndex        =   18
      Top             =   1980
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   4
      Left            =   4560
      TabIndex        =   22
      Top             =   2340
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   5
      Left            =   4560
      TabIndex        =   26
      Top             =   2700
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   6
      Left            =   4560
      TabIndex        =   30
      Top             =   3060
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtCheqDate 
      Height          =   345
      Index           =   7
      Left            =   4560
      TabIndex        =   34
      Top             =   3420
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดเช็คค้างจ้าย"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   4320
      Width           =   1875
   End
   Begin VB.Label lblChequeLeft 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2520
      TabIndex        =   42
      Top             =   4320
      Width           =   1875
   End
   Begin VB.Label lblTotalCheque 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2520
      TabIndex        =   41
      Top             =   3960
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดรวมเช็ค"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   3960
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "วันที่ชำระ"
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   39
      Top             =   210
      Width           =   660
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   -30
      X2              =   8310
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   8340
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "จำนวนเงิน"
      Height          =   195
      Index           =   2
      Left            =   6390
      TabIndex        =   6
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ลงวันที่"
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   4
      Top             =   690
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ธนาคาร"
      Height          =   195
      Index           =   0
      Left            =   2460
      TabIndex        =   2
      Top             =   690
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "เลขที่เช็ค"
      Height          =   195
      Index           =   4
      Left            =   330
      TabIndex        =   1
      Top             =   690
      Width           =   645
   End
End
Attribute VB_Name = "frmCheqPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerFrm_ As Form
Dim curPayCheq_ As BWGCheqPayment
Dim AccNo As String
Dim TotalCheque As Double
Dim TotalPayment As Double

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub TotalPayments(ByVal TotalPay As Double)
    TotalPayment = TotalPay
    txtCheqAmt_Validate 0, False
End Sub

Public Sub SetCallerForm(selFrm As Form)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set callerFrm_ = selFrm
    Set tmpCol = tmpS.CheqPaymentSearch("PayRefID='" & selFrm.RefID & "'")
    If tmpCol.Count > 0 Then
        Set curPayCheq_ = tmpCol(1)
        ShowPayDetail
    Else
        Set curPayCheq_ = New BWGCheqPayment
        curPayCheq_.ID = selFrm.RefID
        curPayCheq_.PayDate = dtPay.ValueYMD
        curPayCheq_.isChanged = False
    End If
End Sub

Public Sub SetAccountNo(ByVal AccountNo As String)
    AccNo = AccountNo
End Sub

Private Sub ShowPayDetail()
    With curPayCheq_
        txtCheqNo(0).Text = .CheqNo1
        txtCheqNo(1).Text = .CheqNo2
        txtCheqNo(2).Text = .CheqNo3
        txtCheqNo(3).Text = .CheqNo4
        txtCheqNo(4).Text = .CheqNo5
        txtCheqNo(5).Text = .CheqNo6
        txtCheqNo(6).Text = .CheqNo7
        txtCheqNo(7).Text = .CheqNo8
        
        txtBankName(0).Text = .CheqBank1
        txtBankName(1).Text = .CheqBank2
        txtBankName(2).Text = .CheqBank3
        txtBankName(3).Text = .CheqBank4
        txtBankName(4).Text = .CheqBank5
        txtBankName(5).Text = .CheqBank6
        txtBankName(6).Text = .CheqBank7
        txtBankName(7).Text = .CheqBank8
        
        If IsDate(.CheqDate1) Then dtCheqDate(0).ValueYMD = .CheqDate1
        If IsDate(.CheqDate2) Then dtCheqDate(1).ValueYMD = .CheqDate2
        If IsDate(.CheqDate3) Then dtCheqDate(2).ValueYMD = .CheqDate3
        If IsDate(.CheqDate4) Then dtCheqDate(3).ValueYMD = .CheqDate4
        If IsDate(.CheqDate5) Then dtCheqDate(4).ValueYMD = .CheqDate5
        If IsDate(.CheqDate6) Then dtCheqDate(5).ValueYMD = .CheqDate6
        If IsDate(.CheqDate7) Then dtCheqDate(6).ValueYMD = .CheqDate7
        If IsDate(.CheqDate8) Then dtCheqDate(7).ValueYMD = .CheqDate8
        
        If .CheqAmt1 > 0 Then txtCheqAmt(0).Text = Format(.CheqAmt1, "#,##0.00")
        If .CheqAmt2 > 0 Then txtCheqAmt(1).Text = Format(.CheqAmt2, "#,##0.00")
        If .CheqAmt3 > 0 Then txtCheqAmt(2).Text = Format(.CheqAmt3, "#,##0.00")
        If .CheqAmt4 > 0 Then txtCheqAmt(3).Text = Format(.CheqAmt4, "#,##0.00")
        If .CheqAmt5 > 0 Then txtCheqAmt(4).Text = Format(.CheqAmt5, "#,##0.00")
        If .CheqAmt6 > 0 Then txtCheqAmt(5).Text = Format(.CheqAmt6, "#,##0.00")
        If .CheqAmt7 > 0 Then txtCheqAmt(6).Text = Format(.CheqAmt7, "#,##0.00")
        If .CheqAmt8 > 0 Then txtCheqAmt(7).Text = Format(.CheqAmt8, "#,##0.00")
    End With
End Sub

Private Sub cmdOK_Click()
    SavePayment
    Unload Me
End Sub

Private Sub dtCheqDate_Validate(Index As Integer, Cancel As Boolean)
Dim cheqDate As String
    If Trim(txtCheqNo(Index).Text) & Trim(txtBankName(Index).Text) & Trim(txtCheqAmt(Index).Text) <> "" Then
        cheqDate = dtCheqDate(Index).ValueYMD
    Else
        cheqDate = ""
    End If
    Select Case Index
        Case 0
            curPayCheq_.CheqDate1 = cheqDate
        Case 1
            curPayCheq_.CheqDate2 = cheqDate
        Case 2
            curPayCheq_.CheqDate3 = cheqDate
        Case 3
            curPayCheq_.CheqDate4 = cheqDate
        Case 4
            curPayCheq_.CheqDate5 = cheqDate
        Case 5
            curPayCheq_.CheqDate6 = cheqDate
        Case 6
            curPayCheq_.CheqDate7 = cheqDate
        Case 7
            curPayCheq_.CheqDate8 = cheqDate
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If curPayCheq_ Is Nothing Then Exit Sub
    If curPayCheq_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            SavePayment
            If curPayCheq_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtBankName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Load frmACList
        frmACList.SetCallerForm Me
        frmACList.Show 1
        If AccNo = "" Then
            If Trim(txtBankName(Index).Text) = "" Then txtBankName(Index).Text = AccNo
        Else
            txtBankName(Index).Text = AccNo
        End If
    ElseIf KeyCode = vbKeyReturn Then
        If Trim(txtBankName(Index).Text) = "" Then Exit Sub
        Load frmACList
        frmACList.SetCallerForm Me
        frmACList.LoadAccount Format(Trim(txtBankName(Index).Text), "00")
        frmACList.SelectedData
        txtBankName(Index).Text = AccNo
    End If
End Sub

Private Sub txtBankName_Validate(Index As Integer, Cancel As Boolean)
Dim cheqDate As String
    If Trim(txtCheqNo(Index).Text) & Trim(txtBankName(Index).Text) & Trim(txtCheqAmt(Index).Text) <> "" Then
        cheqDate = dtCheqDate(Index).ValueYMD
    Else
        cheqDate = ""
    End If
    Select Case Index
        Case 0
            curPayCheq_.CheqBank1 = txtBankName(Index).Text
            curPayCheq_.CheqDate1 = cheqDate
        Case 1
            curPayCheq_.CheqBank2 = txtBankName(Index).Text
            curPayCheq_.CheqDate2 = cheqDate
        Case 2
            curPayCheq_.CheqBank3 = txtBankName(Index).Text
            curPayCheq_.CheqDate3 = cheqDate
        Case 3
            curPayCheq_.CheqBank4 = txtBankName(Index).Text
            curPayCheq_.CheqDate4 = cheqDate
        Case 4
            curPayCheq_.CheqBank5 = txtBankName(Index).Text
            curPayCheq_.CheqDate5 = cheqDate
        Case 5
            curPayCheq_.CheqBank6 = txtBankName(Index).Text
            curPayCheq_.CheqDate6 = cheqDate
        Case 6
            curPayCheq_.CheqBank7 = txtBankName(Index).Text
            curPayCheq_.CheqDate7 = cheqDate
        Case 7
            curPayCheq_.CheqBank8 = txtBankName(Index).Text
            curPayCheq_.CheqDate8 = cheqDate
    End Select
End Sub

Private Sub txtCheqAmt_Validate(Index As Integer, Cancel As Boolean)
Dim cheqDate As String
Dim selAmt As Double
    If IsNumeric(txtCheqAmt(Index).Text) = False Then txtCheqAmt(Index).Text = ""
    If Trim(txtCheqNo(Index).Text) & Trim(txtBankName(Index).Text) & Trim(txtCheqAmt(Index).Text) <> "" Then
        cheqDate = dtCheqDate(Index).ValueYMD
    Else
        cheqDate = ""
    End If
    If IsNumeric(txtCheqAmt(Index).Text) Then
        selAmt = txtCheqAmt(Index).Text
    Else
        selAmt = 0
    End If
    Select Case Index
        Case 0
            curPayCheq_.CheqAmt1 = selAmt
            curPayCheq_.CheqDate1 = cheqDate
        Case 1
            curPayCheq_.CheqAmt2 = selAmt
            curPayCheq_.CheqDate2 = cheqDate
        Case 2
            curPayCheq_.CheqAmt3 = selAmt
            curPayCheq_.CheqDate3 = cheqDate
        Case 3
            curPayCheq_.CheqAmt4 = selAmt
            curPayCheq_.CheqDate4 = cheqDate
        Case 4
            curPayCheq_.CheqAmt5 = selAmt
            curPayCheq_.CheqDate5 = cheqDate
        Case 5
            curPayCheq_.CheqAmt6 = selAmt
            curPayCheq_.CheqDate6 = cheqDate
        Case 6
            curPayCheq_.CheqAmt7 = selAmt
            curPayCheq_.CheqDate7 = cheqDate
        Case 7
            curPayCheq_.CheqAmt8 = selAmt
            curPayCheq_.CheqDate8 = cheqDate
    End Select
    Dim i%
    TotalCheque = 0
    For i = 0 To 7
        TotalCheque = TotalCheque + CDbl(Trim("0" & txtCheqAmt(i)))
    Next i
    lblTotalCheque.Caption = Format(TotalCheque, "#,##0.00")
    lblChequeLeft.Caption = Format(TotalPayment - TotalCheque, "#,##0.00")
End Sub

Private Sub txtCheqNo_Validate(Index As Integer, Cancel As Boolean)
Dim cheqDate As String
    If Trim(txtCheqNo(Index).Text) & Trim(txtBankName(Index).Text) & Trim(txtCheqAmt(Index).Text) <> "" Then
        cheqDate = dtCheqDate(Index).ValueYMD
    Else
        cheqDate = ""
    End If
    Select Case Index
        Case 0
            curPayCheq_.CheqNo1 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate1 = cheqDate
        Case 1
            curPayCheq_.CheqNo2 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate2 = cheqDate
        Case 2
            curPayCheq_.CheqNo3 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate3 = cheqDate
        Case 3
            curPayCheq_.CheqNo4 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate4 = cheqDate
        Case 4
            curPayCheq_.CheqNo5 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate5 = cheqDate
        Case 5
            curPayCheq_.CheqNo6 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate6 = cheqDate
        Case 6
            curPayCheq_.CheqNo7 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate7 = cheqDate
        Case 7
            curPayCheq_.CheqNo8 = txtCheqNo(Index).Text
            curPayCheq_.CheqDate8 = cheqDate
    End Select
End Sub

Private Sub SavePayment()
Dim x%
    For x = 0 To 7
        If Trim(txtCheqNo(x).Text) & Trim(txtBankName(x).Text) & Trim(txtCheqAmt(x).Text) <> "" Then
            If Trim(txtCheqNo(x).Text) = "" Then
                MsgBox "กรุณาระบุเลขที่เช็ค", vbExclamation
                txtCheqNo(x).SetFocus
                Exit Sub
            End If
            If Trim(txtBankName(x).Text) = "" Then
                MsgBox "กรุณาระบุชื่อธนาคารของเช็ค", vbExclamation
                txtBankName(x).SetFocus
                Exit Sub
            End If
            If IsNumeric(txtCheqAmt(x).Text) = False Then
                MsgBox "กรุณาระบุจำนวนเงินในเช็คให้ถูกต้อง", vbExclamation
                txtCheqAmt(x).SetFocus
                Exit Sub
            End If
        End If
    Next
    If curPayCheq_.isChanged Then
        curPayCheq_.Save
    End If
End Sub

