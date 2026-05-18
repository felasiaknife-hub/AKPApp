VERSION 5.00
Begin VB.Form frmImportDataFix 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   285
      Left            =   8430
      TabIndex        =   6
      Top             =   30
      Width           =   1545
   End
   Begin VB.CommandButton Command4 
      Caption         =   "เฉพาะยังไม่มีชื่อย่อ"
      Height          =   345
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "เฉพาะยังไม่มีเขต Sale"
      Height          =   345
      Left            =   1590
      TabIndex        =   3
      Top             =   0
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "แสดงทั้งหมด"
      Height          =   345
      Left            =   5190
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เฉพาะยังไม่มีตำบล"
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin WasteManagment.ctlGrid grItem 
      Height          =   7365
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12991
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   195
      Left            =   6750
      TabIndex        =   5
      Top             =   60
      Width           =   45
   End
End
Attribute VB_Name = "frmImportDataFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colSaleDistrict_ As Collection
Dim curCustID_ As String

Private Sub Command1_Click()
    QueryData "TumbolID is null Or TumbolID=''"
End Sub

Private Sub Command2_Click()
    QueryData ""
End Sub

Private Sub Command3_Click()
    QueryData "SaleDistrictID is null Or SaleDistrictID=''"
End Sub

Private Sub Command4_Click()
    QueryData "CustShortName is null Or CustShortName=''"
End Sub

Private Sub Command5_Click()
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM AA_ExcelImport ORDER BY CustomerID"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
End Sub

Private Sub Form_Load()
    PopulateSaleDistrict
    With grItem.getGridObj
        grItem.ColInputType(11) = ComboBoxEnum
        grItem.ColInputType(12) = ComboBoxEnum
        .Cols = 14
        .TextMatrix(0, 1) = "ชื่อบริษัท"
        .TextMatrix(0, 2) = "ที่อยู่"
        .TextMatrix(0, 3) = "ตำบล"
        .TextMatrix(0, 4) = "อำเภอ"
        .TextMatrix(0, 5) = "จังหวัด"
        .TextMatrix(0, 6) = "รหัสไปรษณีย์"
        .TextMatrix(0, 7) = "โทร."
        .TextMatrix(0, 8) = "แฟกซ์"
        .TextMatrix(0, 9) = "ชื่อผู้ติดต่อ"
        .TextMatrix(0, 10) = "เลขทะเบียนโรงงาน"
        .TextMatrix(0, 11) = "เขตพื้นที่ Sale"
        .TextMatrix(0, 12) = "ใช้งาน?"
        .TextMatrix(0, 13) = "ชื่อย่อ"
        .ColWidth(0) = 0
        .ColWidth(1) = 2355
        .ColWidth(2) = 1695
        .ColWidth(3) = 1395
        .ColWidth(4) = 1215
        .ColWidth(5) = 1080
        .ColWidth(6) = 1125
        .ColWidth(7) = 960
        .ColWidth(8) = 960
        .ColWidth(9) = 960
        .ColWidth(10) = 1530
        .ColWidth(11) = 1245
        .ColWidth(12) = 960
        .ColWidth(13) = 960
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    grItem.Width = Me.ScaleWidth - 60
    grItem.Height = Me.ScaleHeight - 360
End Sub

Private Sub grItem_Click(Row As Integer, Col As Integer)
'    MsgBox grItem.getGridObj.ColWidth(col)
End Sub

Private Sub QueryData(selCriStr As String)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim x%, sqlStr As String
    grItem.ClearAllData
    If Trim(selCriStr) <> "" Then
        sqlStr = "SELECT * FROM tbCustomerImport WHERE " & selCriStr & " ORDER BY CompanyName"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Else
        sqlStr = "SELECT * FROM tbCustomerImport  ORDER BY CompanyName"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    End If
    With grItem.getGridObj
        .Rows = tmpRec.RecordCount + 2
        x = 1
        Do Until tmpRec.EOF
            .TextMatrix(x, 0) = Trim("" & tmpRec("CustomerID"))
            .TextMatrix(x, 1) = Trim("" & tmpRec("CompanyName"))
            .TextMatrix(x, 2) = Trim("" & tmpRec("Addr1"))
            .TextMatrix(x, 3) = Trim("" & tmpRec("TumbolName"))
            .TextMatrix(x, 4) = Trim("" & tmpRec("AumphurName"))
            .TextMatrix(x, 5) = Trim("" & tmpRec("ProvinceName"))
            .TextMatrix(x, 6) = Trim("" & tmpRec("ZipCode"))
            .TextMatrix(x, 7) = Trim("" & tmpRec("TelNo"))
            .TextMatrix(x, 8) = Trim("" & tmpRec("FaxNo"))
            .TextMatrix(x, 9) = Trim("" & tmpRec("ContactPerson"))
            .TextMatrix(x, 10) = Trim("" & tmpRec("RegisterNo"))
            .TextMatrix(x, 11) = Trim("" & tmpRec("SaleDistrictName"))
            .TextMatrix(x, 12) = Trim("" & tmpRec("isUsed"))
            .TextMatrix(x, 13) = Trim("" & tmpRec("CustShortName"))
            tmpRec.MoveNext
            x = x + 1
        Loop
    End With
    Label1.Caption = tmpRec.RecordCount & " รายการ"
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub grItem_KeyDown(KeyCode As Integer, Shift As Integer)
    With grItem.getGridObj
        If .Col = 11 And KeyCode = 46 Then
                DBConnExc "Update tbCustomerImport Set SaleDistrictID='',SaleDistrictName='' WHERE CustomerID='" & Trim(.TextMatrix(.Row, 0)) & "'"
                grItem.getColObject(11).ListIndex = -1
                .TextMatrix(.Row, .Col) = ""
        End If
    End With
End Sub

Private Sub grItem_RolColChanged(Row As Integer, Col As Integer)
Dim x%
Dim tmpCbo As ComboBox
    grItem.ColEnabled(Col) = Not (Col = 4 Or Col = 5)
    If Col = 11 Then
    Set tmpCbo = grItem.getColObject(Col)
    tmpCbo.Clear
        
        For x = 1 To colSaleDistrict_.Count
            tmpCbo.AddItem colSaleDistrict_(x).SaleStaffName & ", " & colSaleDistrict_(x).ZoneName & ", " & colSaleDistrict_(x).DistrictName '& ", " & aa.SaleStaffName
        Next
    ElseIf Col = 12 Then
    Set tmpCbo = grItem.getColObject(Col)
    tmpCbo.Clear
        tmpCbo.AddItem "Y"
        tmpCbo.AddItem "N"
    End If
End Sub

Private Sub grItem_UpdateValue(Row As Integer, Col As Integer)
Dim custID As String
Dim selVal As String
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    custID = Trim(grItem.getGridObj.TextMatrix(Row, 0))
    curCustID_ = custID
    selVal = Trim(grItem.getGridObj.TextMatrix(Row, Col))
    
    Select Case Col
        Case 1
            DBConnExc "Update tbCustomerImport Set CompanyName='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 2
            DBConnExc "Update tbCustomerImport Set Addr1='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 3 'ตำบล
            Set tmpCol = tmpS.TumbolSearch("TumbolName Like '%" & selVal & "%'")
            Dim aa As IMITumbol
            If tmpCol.Count = 1 Then
                DBConnExc "Update tbCustomerImport Set TumbolID='" & tmpCol(1).ID & "',TumbolName='" & tmpCol(1).Name & "',AumphurName='" & tmpCol(1).AumphurName & "',ProvinceName='" & tmpCol(1).ProvinceName & "' WHERE CustomerID='" & custID & "'"
                grItem.getGridObj.TextMatrix(Row, 3) = tmpCol(1).Name
                grItem.getGridObj.TextMatrix(Row, 4) = tmpCol(1).AumphurName
                grItem.getGridObj.TextMatrix(Row, 5) = tmpCol(1).ProvinceName
            ElseIf tmpCol.Count > 1 Then
                Load frmTumbolSearch
                With frmTumbolSearch
                    .SetCallerForm Me
                    .getCriteria selVal
                    '.txtCriteria.Text = selVal
                    .Show 1
                End With
            ElseIf tmpCol.Count = 0 Then
                DBConnExc "Update tbCustomerImport Set TumbolID='',TumbolName='',AumphurName='',ProvinceName='' WHERE CustomerID='" & custID & "'"
                grItem.getGridObj.TextMatrix(Row, 3) = ""
                grItem.getGridObj.TextMatrix(Row, 4) = ""
                grItem.getGridObj.TextMatrix(Row, 5) = ""
            End If
        Case 6
            DBConnExc "Update tbCustomerImport Set ZipCode='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 7
            DBConnExc "Update tbCustomerImport Set TelNo='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 8
            DBConnExc "Update tbCustomerImport Set FaxNo='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 9
            DBConnExc "Update tbCustomerImport Set ContactPerson='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 10
            DBConnExc "Update tbCustomerImport Set RegisterNo='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 11 'เขตพื้นที่
            If Trim(selVal) <> "" Then
                DBConnExc "Update tbCustomerImport Set SaleDistrictID='" & colSaleDistrict_(grItem.getColObject(Col).ListIndex + 1).ID & "',SaleDistrictName='" & grItem.getColObject(Col).Text & "' WHERE CustomerID='" & custID & "'"
            End If
        Case 12
            DBConnExc "Update tbCustomerImport Set isUsed='" & selVal & "' WHERE CustomerID='" & custID & "'"
        Case 13
            DBConnExc "Update tbCustomerImport Set CustShortName='" & selVal & "' WHERE CustomerID='" & custID & "'"
    End Select

End Sub

Private Sub PopulateSaleDistrict()
Dim tmpS As New BWGSearchManager
    Set colSaleDistrict_ = tmpS.SaleDistrictSearch
    Set tmpS = Nothing
End Sub

Public Sub SetDataTumbol(selTumbol As IMITumbol)
    DBConnExc "Update tbCustomerImport Set TumbolID='" & selTumbol.ID & "',TumbolName='" & selTumbol.Name & "',AumphurName='" & selTumbol.AumphurName & "',ProvinceName='" & selTumbol.ProvinceName & "' WHERE CustomerID='" & curCustID_ & "'"
    With grItem.getGridObj
        .TextMatrix(.Row, 3) = selTumbol.Name
        .TextMatrix(.Row, 4) = selTumbol.AumphurName
        .TextMatrix(.Row, 5) = selTumbol.ProvinceName
    End With
End Sub
