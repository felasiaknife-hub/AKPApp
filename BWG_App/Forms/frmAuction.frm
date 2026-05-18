VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmAuction 
   Caption         =   "เพิ่มข้อมูลงานประมูล"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   1530
      TabIndex        =   19
      Top             =   1890
      Width           =   2895
      Begin VB.OptionButton IsNoK 
         Caption         =   "ไม่ได้งาน"
         Height          =   285
         Left            =   1350
         TabIndex        =   22
         Top             =   135
         Width           =   1140
      End
      Begin VB.OptionButton IsOK 
         Caption         =   "ได้งาน"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.TextBox txtPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1530
      TabIndex        =   17
      Top             =   1575
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1005
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   7095
   End
   Begin VB.TextBox txtRemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   1215
      Width           =   7095
   End
   Begin VB.TextBox txtContractNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   435
      Width           =   3540
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   765
      TabIndex        =   0
      Top             =   3585
      Visible         =   0   'False
      Width           =   4245
   End
   Begin DateCtl.Date dtStart 
      Height          =   315
      Left            =   1530
      TabIndex        =   2
      Top             =   825
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin DateCtl.Date dtEnd 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   2400
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin DateCtl.Date dtNext 
      Height          =   315
      Left            =   5175
      TabIndex        =   23
      Top             =   2430
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ประมูลครั้งต่อไป"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   24
      Top             =   2490
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   1215
      TabIndex        =   20
      Top             =   2070
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   18
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label lblAuctionID 
      Height          =   105
      Left            =   7965
      TabIndex        =   16
      Top             =   765
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   870
      TabIndex        =   15
      Top             =   75
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   810
      TabIndex        =   14
      Top             =   1275
      Width           =   675
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   0
      Top             =   2775
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ประมูล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   13
      Top             =   885
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่สิ้นสุดสัญญา"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2460
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่งานประมูล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   390
      TabIndex        =   11
      Top             =   495
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บริษัท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   270
      TabIndex        =   10
      Top             =   3645
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "frmAuction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myAuction_ As BWGCustomerAuction
Dim curCustomer_ As BWGCustomer
Dim curCompany_ As BWGCompanyConfig
Dim callerForm_ As Form
Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลงานประมูลของลูกค้ารายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myAuction_.Delete(Trim(lblAuctionID.Caption))
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewAuction
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myAuction_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myAuction_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewAuction
'    txtCompany.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If

'    If Trim(txtCompany.Text) = "" Then
'        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
'        txtCompany.SetFocus
'        Exit Sub
'    End If
'
    If Trim(txtContractNo.Text) = "" Then
        MsgBox "กรุณาระบุเลขที่การประมูล", vbExclamation
        txtContractNo.SetFocus
        Exit Sub
    End If
'    If dtStart.ValueYMD = dtEnd.ValueYMD Then
'        MsgBox "ข้อมูลวันเริ่มต้นและสิ้นสุดการประมูลไม่ถูกต้อง จะต้องไม่เป็นวันเดียวกัน", vbExclamation
'        dtEnd.SetFocus
'        Exit Sub
'    End If
    myAuction_.docNo = Trim(txtContractNo.Text)
    myAuction_.StartDate = Trim(dtStart.ValueYMD)
    myAuction_.Remdesc = Trim(txtRemDesc.Text)
    myAuction_.Price = Trim(txtPrice.Text)
    myAuction_.NextDate = Trim(dtNext.ValueYMD)
    If IsOK.Value = True Then
        myAuction_.IsOK = "Y"
        myAuction_.EndDate = Trim(dtEnd.ValueYMD)
    End If
    If IsNoK.Value = True Then
        myAuction_.IsOK = "N"
        myAuction_.EndDate = ""
    End If
    If myAuction_.isChanged Then
        myAuction_.Save (Trim(lblAuctionID.Caption))
    End If
    If myAuction_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub dtEnd_Validate(Cancel As Boolean)
    myAuction_.EndDate = Trim(dtEnd.ValueYMD)
End Sub

Private Sub dtStart_Validate(Cancel As Boolean)
    myAuction_.StartDate = Trim(dtStart.ValueYMD)
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    CreateNewAuction
    If IsOK.Value = True Then
        dtEnd.Enabled = True
    Else
        dtEnd.Enabled = False
    End If
    Set curCompany_ = New BWGCompanyConfig
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtContractNo.Text = ""
    dtStart.ValueYMD = TodayDate
    dtEnd.ValueYMD = TodayDate
    dtNext.ValueYMD = TodayDate
    txtRemDesc.Text = ""
    lblAuctionID.Caption = ""
End Sub

Private Sub CreateNewAuction()
    Set myAuction_ = Nothing
    Set myAuction_ = New BWGCustomerAuction
    With myAuction_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        If Not curCompany_ Is Nothing Then
            .CompanyID = curCompany_.ID
            .CompanyName = curCompany_.CompanyNameTH
        End If
        .StartDate = dtStart.ValueYMD
        .isChanged = False
    End With
End Sub
Public Sub SetDataAuction(selAuction As BWGCustomerAuction)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myAuction_ = Nothing
    Set myAuction_ = selAuction
    If myAuction_ Is Nothing Then Exit Sub
    With myAuction_
        Set tmpCol = tmpS.CompanySearch("CompanyID='" & .CompanyID & "'")
        If tmpCol.Count > 0 Then SetDataCompany tmpCol(1), False
        txtContractNo.Text = .docNo
        dtStart.ValueYMD = .StartDate
        If .EndDate <> "" Then
            dtEnd.ValueYMD = .EndDate
        End If
        txtRemDesc.Text = .Remdesc
        txtPrice.Text = .Price
        dtNext.ValueYMD = .NextDate
        If .IsOK = "Y" Then
            IsOK.Value = True
        ElseIf .IsOK = "N" Then
            IsNoK.Value = True
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myAuction_ Is Nothing Then Exit Sub
    If myAuction_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myAuction_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetDataCompany(selCompany As BWGCompanyConfig, Optional isAssigned As Boolean = True)
    Set curCompany_ = Nothing
    Set curCompany_ = selCompany
    If curCompany_ Is Nothing Then Exit Sub
    With curCompany_
        txtCompany.Text = .CompanyNameTH
        If isAssigned Then
            myAuction_.CompanyID = .ID
            myAuction_.CompanyName = .CompanyNameTH
        End If
    End With
End Sub

Public Sub SetCallerForm(selFrm As Form, selCust As BWGCustomer)
    Set callerForm_ = selFrm
    Set curCustomer_ = selCust
    txtCustomer.Text = selCust.CustomerName
    If Not myAuction_ Is Nothing Then myAuction_.CustomerID = curCustomer_.ID
End Sub
Private Sub IsNoK_Click()
    dtEnd.Enabled = False
End Sub
Private Sub IsOK_Click()
    dtEnd.Enabled = True
End Sub
Private Sub txtCompany_Validate(Cancel As Boolean)
    If Trim(txtCompany.Text) = "" And Trim(curCompany_.CompanyNameTH) <> Trim(txtCompany.Text) Then
        Set curCompany_ = Nothing
        Set curCompany_ = New BWGCompanyConfig
        With myAuction_
            .CompanyID = ""
            .CompanyName = ""
        End With
    Else
        If Trim(curCompany_.CompanyNameTH) <> Trim(txtCompany.Text) Then
            Set curCompany_ = Nothing
            Set curCompany_ = New BWGCompanyConfig
            With myAuction_
                .CompanyID = ""
                .CompanyName = ""
            End With
            Load frmCompanySearch
            With frmCompanySearch
                .getCriteria (txtCompany.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myAuction_.CompanyID = "" Then txtCompany.Text = ""
        End If
    End If
End Sub

Private Sub txtContractNo_Validate(Cancel As Boolean)
    myAuction_.docNo = Trim(txtContractNo.Text)
End Sub

Private Sub txtRemDesc_Validate(Cancel As Boolean)
    myAuction_.Remdesc = Trim(txtRemDesc.Text)
End Sub

