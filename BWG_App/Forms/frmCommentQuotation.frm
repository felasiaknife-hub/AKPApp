VERSION 5.00
Begin VB.Form frmCommentQuo 
   Caption         =   "บันทึกข้อตกลงการให้บริการ "
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "เพิ่มใบเสนอราคา"
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   1005
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   7095
   End
   Begin VB.TextBox txtRemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   7095
   End
   Begin VB.TextBox txtContractNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   3555
   End
   Begin VB.TextBox txtCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   4515
   End
   Begin WasteManagment.ctlDate dtStart 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1500
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtEnd 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   1500
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlGrid grItem 
      Height          =   2295
      Left            =   1080
      TabIndex        =   12
      Top             =   2400
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4048
   End
   Begin VB.Label lbcmmID 
      Caption         =   "."
      Height          =   135
      Left            =   1800
      TabIndex        =   18
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   1020
      TabIndex        =   11
      Top             =   270
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   840
      TabIndex        =   10
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เริ่ม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่สิ้นสุด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3630
      TabIndex        =   8
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่บันทึกข้อตกลง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตกลงกับบริษัท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   360
      TabIndex        =   6
      Top             =   690
      Width           =   1005
   End
End
Attribute VB_Name = "frmCommentQuo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myContract_ As BWGContract
Dim curCustomer_ As BWGCustomer
Dim curCompany_ As BWGCompanyConfig
Dim myCmmQuoHeader_ As BWGCmmQuoHeader
Dim myCmmQuoDetail_ As BWGCmmQuoDetail
Dim callerForm_ As Form
Dim i As Integer

Private Sub cmdCancel_Click()
    DBConnExc "Delete tbCommentQuoHeader WHERE CmID='" & lbcmmID.Caption & "' "
    DBConnExc "Delete tbCommentQuoDetail WHERE CmID='" & lbcmmID.Caption & "'"
    MsgBox "ลบข้อมูลเรียบร้อยแล้ว"
    cmdClear_Click
End Sub

Private Sub cmdClear_Click()
    grItem.ClearAllData
    Call SetMSFlexGrid
    lbcmmID.Caption = ""
    txtCompany.Text = ""
    txtContractNo.Text = ""
    txtRemDesc.Text = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Set myCmmQuoDetail_ = New BWGCmmQuoDetail
    If Trim(txtContractNo.Text) = "" Then
        MsgBox "กรุณาระบุเลบที่บันทึกข้อตกลง เพื่อใช้ในการตรวจสอบ"
        Exit Sub
    End If
    If lbcmmID.Caption = "" Or lbcmmID.Caption = "." Then
        lbcmmID.Caption = GetGUID
    End If
    myCmmQuoHeader_.ID = lbcmmID.Caption
    myCmmQuoHeader_.CustomerID = curCustomer_.ID
    myCmmQuoHeader_.CompanyID = curCompany_.ID
    myCmmQuoHeader_.CmQuoNo = Trim(txtContractNo.Text)
    myCmmQuoHeader_.StartDate = dtStart.ValueYMD
    myCmmQuoHeader_.EndDate = dtEnd.ValueYMD
    myCmmQuoHeader_.Comment = Trim(txtRemDesc.Text)
    myCmmQuoHeader_.Save
    DBConnExc "Delete tbCommentQuoDetail WHERE CmID='" & myCmmQuoHeader_.ID & "'"
    Dim i As Integer
    With grItem.getGridObj
        For i = 1 To grItem.getGridObj.Rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Then
                .TextMatrix(i, 0) = i
                myCmmQuoDetail_.CmID = myCmmQuoHeader_.ID
                myCmmQuoDetail_.QuotationID = .TextMatrix(i, 1)
                myCmmQuoDetail_.QuotationNo = .TextMatrix(i, 2)
                myCmmQuoDetail_.isDateS = .TextMatrix(i, 3)
                myCmmQuoDetail_.QuoStatus = .TextMatrix(i, 4)
                If myCmmQuoDetail_.QuotationID = "" Then
                    .TextMatrix(i, 0) = ""
                Else
                    myCmmQuoDetail_.ID = GetGUID
                    myCmmQuoDetail_.Save
                End If
            End If
        Next i
    End With
    MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว"
End Sub

Private Sub Command1_Click()
  If Trim(txtCustomer.Text) = "" Or txtCompany.Text = "" Then
        MsgBox "กรุณาระบุบริษัทลูกค้า และ บริษัท"
        Exit Sub
    End If
   Load frmQuotationSearch
   With frmQuotationSearch
        .getCriteria (txtCustomer.Text)
        .SetCallerForm Me
        .Show 1
   End With
End Sub

Private Sub Form_Load()
    cmdClear_Click
End Sub
Public Sub SetCallerForm(selFrm As Form, selCust As BWGCustomer)
    Set callerForm_ = selFrm
    Set curCustomer_ = selCust
    txtCustomer.Text = selCust.CustomerName
    If Not myContract_ Is Nothing Then myContract_.CustomerID = curCustomer_.ID
End Sub
Public Sub SetDataQuotation(selItem As BWGQuotationHeader)
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim criStr As String
Dim result As Integer
Dim strCond As String
Dim x%, i%, k%
Dim tmpList As ListItem
Dim myQuotation_ As BWGQuotationHeader
Dim myCmmQuoHeader_ As BWGCmmQuoHeader
Set myCmmQuoDetail_ = New BWGCmmQuoDetail
    Set myQuotation_ = Nothing
    Set myQuotation_ = selItem
     If myQuotation_ Is Nothing Then Exit Sub
    With myQuotation_
        myCmmQuoDetail_.QuotationID = .ID
        myCmmQuoDetail_.QuotationNo = .QuotationNo
        myCmmQuoDetail_.isDateS = .IssuedDate
        If .isClosed = True Then
            myCmmQuoDetail_.QuoStatus = "Closed"
        ElseIf .isCancel = True Then
            myCmmQuoDetail_.QuoStatus = "ยกเลิก"
        ElseIf .IsApproved = True Then
            myCmmQuoDetail_.QuoStatus = "อนุมัติแล้ว"
        ElseIf .IsApproved = False Then
            myCmmQuoDetail_.QuoStatus = "ยังไม่อนุมัติ"
        Else
            myCmmQuoDetail_.QuoStatus = "ยังไม่อนุมัติ"
        End If
    End With
    Dim Duplicate As Boolean
    Duplicate = False
    With grItem.getGridObj
        For i = 1 To grItem.getGridObj.Rows - 1
            If Trim(.TextMatrix(i, 0)) = "" Then
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = myCmmQuoDetail_.QuotationID
                .TextMatrix(i, 2) = myCmmQuoDetail_.QuotationNo
                .TextMatrix(i, 3) = myCmmQuoDetail_.isDateS
                .TextMatrix(i, 4) = myCmmQuoDetail_.QuoStatus
                .Rows = .Rows + 1
                Exit For
            Else
                For k = 2 To .Rows - 1
                    If myCmmQuoDetail_.QuotationID = Trim(.TextMatrix(k - 1, 1)) Then Duplicate = True: Exit For
                Next k
                If Duplicate = True Then Exit For
            End If
        Next i
    End With
End Sub
Public Sub SetDataCmmQuo(selItem As BWGCmmQuoHeader)
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim criStr As String
Dim result As Integer
Dim strCond As String
Dim x%
Dim tmpList As ListItem
Dim myQuotation_ As BWGQuotationHeader
Dim myCmmQuoHeader_ As BWGCmmQuoHeader
Set myCmmQuoDetail_ = New BWGCmmQuoDetail
    Set myQuotation_ = Nothing
    Set myQuotation_ = selItem
     If myQuotation_ Is Nothing Then Exit Sub
    With myQuotation_
        myCmmQuoDetail_.QuotationID = .ID
        myCmmQuoDetail_.QuotationNo = .QuotationNo
        myCmmQuoDetail_.isDateS = .IssuedDate
        If .isClosed = True Then
            myCmmQuoDetail_.QuoStatus = "Closed"
        ElseIf .isCancel = True Then
            myCmmQuoDetail_.QuoStatus = "ยกเลิก"
        ElseIf .IsApproved = True Then
            myCmmQuoDetail_.QuoStatus = "อนุมัติแล้ว"
        ElseIf .IsApproved = False Then
            myCmmQuoDetail_.QuoStatus = "ยังไม่อนุมัติ"
        Else
            myCmmQuoDetail_.QuoStatus = "ยังไม่อนุมัติ"
        End If
    End With
    With grItem.getGridObj
        If i = 0 Then i = 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = myCmmQuoDetail_.QuotationID
            .TextMatrix(i, 2) = myCmmQuoDetail_.QuotationNo
            .TextMatrix(i, 3) = myCmmQuoDetail_.isDateS
            .TextMatrix(i, 4) = myCmmQuoDetail_.QuoStatus
    End With
End Sub
Public Sub SetDataCmmQuotation(selItem As BWGCmmQuoHeader)
On Error Resume Next
Call SetMSFlexGrid
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myCmmQuoHeader_ = Nothing
    Set myCmmQuoHeader_ = selItem
    If myCmmQuoHeader_ Is Nothing Then Exit Sub
    With myCmmQuoHeader_
            lbcmmID.Caption = .ID
           txtContractNo.Text = .CmQuoNo
           .CompanyID = curCompany_.ID
           txtCompany.Text = .CompanyNameTH
            dtStart.ValueYMD = .StartDate
            dtEnd.ValueYMD = .EndDate
            txtRemDesc.Text = .Comment
    End With
    ShowQuoCmmDetails
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub
Private Sub ShowQuoCmmDetails()
Dim x%, Y%
Dim tmpCol As Collection
Dim tmpDet As BWGCmmQuoDetail
Dim tmpKey As String, rowCount As Long
Dim forColor As String
Dim HeatValue As Double
Dim gr As ctlGrid
Dim i As Integer
    Set tmpCol = myCmmQuoHeader_.QuotationCmmDetails
        With grItem.getGridObj
            grItem.ClearAllData
            rowCount = 0
            forColor = vbBlack
            If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
            For x = 1 To tmpCol.Count
                Set tmpDet = tmpCol(x)
                .TextMatrix(x + rowCount, 0) = x
                .Row = x + rowCount
                .Col = 1
                .CellForeColor = forColor
                .TextMatrix(x + rowCount, 1) = tmpDet.ID
                .TextMatrix(x + rowCount, 2) = tmpDet.QuotationNo
                .TextMatrix(x + rowCount, 3) = tmpDet.isDateS
                .TextMatrix(x + rowCount, 4) = tmpDet.QuoStatus
            Next
            .Row = 1
            .Col = 1
            .Rows = .Rows + 5
        End With
    Set tmpCol = Nothing
End Sub

Private Sub SetMSFlexGrid()
On Error Resume Next
    With grItem.getGridObj
        .Cols = 5
        .ColWidth(0) = 400
        .ColWidth(1) = 0
        .ColWidth(2) = 3540
        .ColWidth(3) = 1110
        .ColWidth(4) = 1100
        .Rows = 9
        .FixedAlignment(0) = 4
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .TextMatrix(0, 0) = "ที่"
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "ใบเสนอรา"
        .TextMatrix(0, 3) = "วันที่"
        .TextMatrix(0, 4) = "สถานะ"
    End With
End Sub

 Private Sub grItem_Click(Row As Integer, Col As Integer)
    If Row = 0 Then
        i = 1
    Else
        i = Row
    End If
End Sub

Private Sub txtCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Public Sub SetDataCompany(selCompany As BWGCompanyConfig, Optional isAssigned As Boolean = True)
    Set curCompany_ = Nothing
    Set curCompany_ = selCompany
    Set myCmmQuoHeader_ = New BWGCmmQuoHeader
    If curCompany_ Is Nothing Then Exit Sub
    With curCompany_
        txtCompany.Text = .CompanyNameTH
        myCmmQuoHeader_.CompanyID = .ID
    End With
End Sub

Private Sub txtCompany_Validate(Cancel As Boolean)
            Load frmCompanySearch
            With frmCompanySearch
                .getCriteria (txtCompany.Text)
                .SetCallerForm Me
                .Show 1
            End With
End Sub

Private Sub txtContractNo_Validate(Cancel As Boolean)
    myCmmQuoHeader_.CmQuoNo = Trim(txtContractNo.Text)
End Sub

Private Sub txtRemDesc_Validate(Cancel As Boolean)
    myCmmQuoHeader_.CmQuoNo = Trim(txtRemDesc.Text)
End Sub
