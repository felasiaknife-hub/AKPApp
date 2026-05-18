VERSION 5.00
Begin VB.Form frmQuotationDiscout 
   Caption         =   "ส่วนลดใบเสนอราคา"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   615
      Left            =   4140
      TabIndex        =   23
      Top             =   6720
      Width           =   975
   End
   Begin WasteManagment.ctlGrid grItem 
      Height          =   2295
      Left            =   60
      TabIndex        =   16
      Top             =   4320
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ปิด"
      Height          =   615
      Left            =   7380
      TabIndex        =   15
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   615
      Left            =   5220
      TabIndex        =   14
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "สร้างใหม่"
      Height          =   615
      Left            =   6300
      TabIndex        =   13
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "บันทึก"
      Height          =   615
      Left            =   3060
      TabIndex        =   12
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton CmdQuotationAdd 
      Caption         =   "เพิ่มใบเสนอราคา"
      Height          =   615
      Left            =   60
      TabIndex        =   11
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "ส่วนลด"
      Height          =   2655
      Left            =   60
      TabIndex        =   5
      Top             =   1440
      Width           =   8295
      Begin VB.TextBox txtComment 
         BackColor       =   &H80000014&
         Height          =   975
         Left            =   1440
         TabIndex        =   9
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox txtDiscount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   3000
         TabIndex        =   20
         Top             =   1080
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   5640
         TabIndex        =   21
         Top             =   1080
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6300
         TabIndex        =   24
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   22
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ใบเสนอราคานี้ใช้ได้ช่วงวันที่ :"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lbQuoDisNo 
         BackColor       =   &H00C0E0FF&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "เลขที่ :"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "หมายเหตุ :"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "เปอร์เซน"
         Height          =   255
         Left            =   3720
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "ใบส่วนลดนี้สามารถลดราคาได้ :"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1140
      TabIndex        =   1
      Top             =   1005
      Width           =   7215
   End
   Begin WasteManagment.ctlDate dtIssuedDate 
      Height          =   315
      Left            =   6540
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "วันที่ :"
      Height          =   255
      Left            =   5820
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "ใบส่วนลดใบเสนอราคา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   -2400
      X2              =   10560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   4920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   0
      Picture         =   "frmQuotationDiscout.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Label Label1 
      Caption         =   "บริษัทลูกค้า :"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmQuotationDiscout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myQuotationDisCount_ As New BWGQTDiscountHeader
Dim myQuotationDisDetail_ As New BWGQTDiscountDetail
Dim curCustomer_ As BWGCustomer
Dim isDetailChange_ As Boolean
Dim colSearchResult_ As Collection
Dim i As Integer
Public isQuotationDiscount As Boolean


Private Sub cmdCancel_Click()
    If myQuotationDisCount_ Is Nothing Then Exit Sub
    If MsgBox("", vbOKCancel, "") = vbCancel Then
        Exit Sub
    End If
    myQuotationDisCount_.CancelQuotationDiscount
End Sub

Private Sub CmdQuotationAdd_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุบริษัทลูกค้า "
        Exit Sub
    End If
    isQuotationDiscount = True
   Load frmQuotationSearch
   With frmQuotationSearch
        .getCriteria (txtCustomer.Text)
        .SetCallerForm Me
        .Show 1
   End With
End Sub

Private Sub cmdSave_Click()
    If myQuotationDisCount_.CustomerID = "" Or Trim(txtCustomer.Text) = "" Then
        MsgBox ("กรุณาระบุชื่อบริษัทลูกค้า ก่อน")
        Exit Sub
    End If
         With grItem.getGridObj
               If .TextMatrix(1, 1) = "" Then
                    MsgBox ("กรุณาระบุใบเสนอราคาอย่างน้อย 1 รายการ")
                Exit Sub
               End If
        End With
    If myQuotationDisCount_.ID = "" Then
        myQuotationDisCount_.ID = GetGUID
    End If
     myQuotationDisCount_.StartDate = dtFrom.ValueYMD
     myQuotationDisCount_.EndDate = dtTo.ValueYMD
     myQuotationDisCount_.Save
     Dim j As Integer
     DBConnExc "Delete tbQuotationDiscountDetail WHERE QuoDiscountHeaderID='" & myQuotationDisCount_.ID & "'"
     For j = 1 To 8
         With grItem.getGridObj
                myQuotationDisDetail_.QuoDisCountHeaderID = myQuotationDisCount_.ID
                .TextMatrix(j, 0) = j
                myQuotationDisDetail_.QuotationID = .TextMatrix(j, 1)
                myQuotationDisDetail_.QuotationNo = .TextMatrix(j, 2)
                myQuotationDisDetail_.QuotationDate = .TextMatrix(j, 3)
                myQuotationDisDetail_.QuotationStatus = .TextMatrix(j, 4)
        End With
        If myQuotationDisDetail_.QuotationID = "" Then
            With grItem.getGridObj
                .TextMatrix(j, 0) = ""
            End With
            Exit For
        Else
            myQuotationDisDetail_.ID = GetGUID
            myQuotationDisDetail_.Save
            DBConnExc "Update tbQuotationHeader Set isQuoDisount = 'Y' WHERE QuotationID='" & myQuotationDisDetail_.QuotationID & "'"
        End If
     Next
     lbQuoDisNo.Caption = myQuotationDisCount_.QuoDisNo
     MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว"
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGQTDiscountHeader
Dim tmpCol As Collection
    Load frmQuotationDiscountSearch
    With frmQuotationDiscountSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Command2_Click()
myQuotationDisCount_.ID = ""
txtCustomer.Text = ""
lbQuoDisNo.Caption = "*"
txtComment.Text = ""
txtDiscount.Text = "0"
grItem.ClearAllData
SetMSFlexGrid
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{Tab}"
End Sub


Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Call SetMSFlexGrid
    isQuotationDiscount = False
End Sub

Private Sub grItem_Click(Row As Integer, Col As Integer)
    If Row = 0 Then
        i = 1
    Else
        i = Row
    End If
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myQuotationDisCount_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myQuotationDisCount_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myQuotationDisCount_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub
Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        Set tmpRec = Nothing
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myQuotationDisCount_.CustomerID = .ID
            myQuotationDisCount_.CustomerName = .CustomerName
            myQuotationDisCount_.QuoDisNo = .CustomerCode & "/" & dtIssuedDate.YearValue & dtIssuedDate.MonthValue & dtIssuedDate.DayValue
        End If
    End With
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub
Public Sub SetDataQuotationDiscount(selItem As BWGQTDiscountHeader)
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myQuotationDisCount_ = Nothing
    Set myQuotationDisCount_ = selItem
    If myQuotationDisCount_ Is Nothing Then Exit Sub
    With myQuotationDisCount_
            txtCustomer.Text = .CustomerName
            txtDiscount.Text = .Discount
            txtComment.Text = .Detail
            lbQuoDisNo.Caption = .QuoDisNo
            dtFrom.ValueYMD = .StartDate
            dtTo.ValueYMD = .EndDate
    End With
    ShowQuoDiscountDetails
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub
Private Sub ShowQuoDiscountDetails()
Dim x%, Y%
Dim tmpCol As Collection
Dim tmpDet As BWGQTDiscountDetail
Dim tmpKey As String, rowCount As Long
Dim forColor As String
Dim HeatValue As Double
Dim gr As ctlGrid
Dim i As Integer
    Set tmpCol = myQuotationDisCount_.QuotationDiscountDetails
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
                .TextMatrix(x + rowCount, 3) = tmpDet.QuotationDate
                .TextMatrix(x + rowCount, 4) = tmpDet.QuotationStatus
            Next
            .Row = 1
            .Col = 1
        End With
    Set tmpCol = Nothing
End Sub
Public Sub SetDataQuotation(selItem As BWGQuotationHeader)
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim criStr As String
Dim result As Integer
Dim strCond As String
Dim x%
Dim tmpList As ListItem
Dim myQuotation_ As BWGQuotationHeader
Dim myQuotationDisH_ As BWGQTDiscountHeader
Set myQuotationDisDetail_ = New BWGQTDiscountDetail
    Set myQuotation_ = Nothing
    Set myQuotation_ = selItem
     If myQuotation_ Is Nothing Then Exit Sub
    With myQuotation_
        myQuotationDisDetail_.QuotationID = .ID
        myQuotationDisDetail_.QuotationNo = .QuotationNo
        myQuotationDisDetail_.QuotationDate = .IssuedDate
        If .isClosed = True Then
            myQuotationDisDetail_.QuotationStatus = "Closed"
        ElseIf .isCancel = True Then
            myQuotationDisDetail_.QuotationStatus = "ยกเลิก"
        ElseIf .IsApproved = True Then
            myQuotationDisDetail_.QuotationStatus = "อนุมัติแล้ว"
        ElseIf .IsApproved = False Then
            myQuotationDisDetail_.QuotationStatus = "ยังไม่อนุมัติ"
        Else
            myQuotationDisDetail_.QuotationStatus = "ยังไม่อนุมัติ"
        End If
    End With
    With grItem.getGridObj
        If i = 0 Then i = 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = myQuotationDisDetail_.QuotationID
            .TextMatrix(i, 2) = myQuotationDisDetail_.QuotationNo
            .TextMatrix(i, 3) = myQuotationDisDetail_.QuotationDate
            .TextMatrix(i, 4) = myQuotationDisDetail_.QuotationStatus
    End With
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


Private Sub txtDiscount_Validate(Cancel As Boolean)
    myQuotationDisCount_.Discount = Trim(txtDiscount.Text)
End Sub

Private Sub txtComment_Validate(Cancel As Boolean)
    myQuotationDisCount_.Detail = Trim(txtComment.Text)
End Sub
