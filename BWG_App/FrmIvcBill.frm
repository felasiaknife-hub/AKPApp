VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmIvcBill 
   Caption         =   "ใบรับวางบิล"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DtpFromDate 
      Height          =   330
      Left            =   1800
      TabIndex        =   15
      Top             =   1845
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   96272385
      CurrentDate     =   43654
   End
   Begin Crystal.CrystalReport Rpt1 
      Left            =   135
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox TxtCustomer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Text            =   "*"
      Top             =   945
      Width           =   3645
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "*"
      Top             =   1290
      Width           =   3645
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1710
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "พิมพ์ใบวางบิล"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1290
      Width           =   2595
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   870
      Width           =   2595
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2730
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบวางบิล"
         Object.Width           =   3334
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่"
         Object.Width           =   2937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "ยอดใบวางบิล (ก่อน VAT)"
         Object.Width           =   3360
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ยอดที่ชำระแล้ว"
         Object.Width           =   3307
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   5670
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบวางบิล"
         Object.Width           =   2805
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่"
         Object.Width           =   2761
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "ยอดใบวางบิล (ก่อน VAT)"
         Object.Width           =   3713
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vat"
         Object.Width           =   2778
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "ยอดรวม Vat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เลขบิล"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpTodate 
      Height          =   330
      Left            =   4005
      TabIndex        =   18
      Top             =   1845
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   96272385
      CurrentDate     =   43654
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   4
      Left            =   3375
      TabIndex        =   17
      Top             =   1890
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   3
      Left            =   1395
      TabIndex        =   16
      Top             =   1890
      Width           =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   3
      X1              =   225
      X2              =   7035
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   315
      X2              =   7125
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกใบวางบิล"
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1380
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "FrmIvcBill.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกใบวางบิล"
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
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   90
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   11
      Top             =   1305
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ผลการค้นหา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   2490
      Width           =   8655
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   8700
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "รายการสินค้าที่เลือก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   5430
      Width           =   8655
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   8700
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   8
      Top             =   990
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบวางบิล"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   1605
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดเงินที่แสดงในหน้าจอนี้เป็นยอดรวม VAT และหักภาษี ณ. ที่จ่ายแล้ว"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   2220
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "FrmIvcBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim curBill_ As BWGRepIvcBill
Dim colOriSelProduct_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim callerFrm_ As Form
Dim invHeaderID_ As String
Dim companyID_ As String

Private Sub cmdClose_Click()
    curBill_.DeleteAll
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim x%
Dim tmpKey As String
    If lvSelected.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูล กรุณาตรวจสอบว่าได้ทำการเลือกข้อมูลแล้ว", vbOKOnly + vbExclamation, "กรุณาตรวจสอบ"
        Exit Sub
    End If
        PrintMain 0
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, iCount As Long
    tmpSql = ""
    tmpSql = "Select * from TbTmpRepIvcBill" & vbCrLf
    tmpSql = tmpSql & " WHERE UserN='" & CurrentUser.LoginName & "' "
    With rpt1
        .Reset
        .WindowTitle = "ใบรับวางบิล"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = False
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = False
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\RepIvcBill.rpt"
        .SQLQuery = tmpSql
        .Formulas(1) = "StaffName=" & Chr(34) & CurrentUser.StaffFName & " " & CurrentUser.StaffLName & Chr(34)
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    curBill_.DeleteAll
End Sub
Private Sub cmdSearch_Click()
Dim tmpS As New BWGSearchManager
Dim tmpList As ListItem
Dim tmpBill As BWGBillingHeader
Dim tmpArr
Dim criStr1 As String, tmpStr$, curNameSpec As String
Dim sqlStr$
Dim x%
    If Trim(txtKey.Text) = "" Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtKey.SetFocus
        Exit Sub
    End If
    criStr1 = Replace(Trim(txtKey.Text), "*", "%")
    sqlStr = "(BillingNo Like '%" & criStr1 & "%')"
    If Not curCustomer_ Is Nothing Then
        sqlStr = sqlStr & " And CustomerID='" & curCustomer_.ID & "'"
    End If
    sqlStr = sqlStr & " and IssuedDate>='" & FormatDMY_to_YMD(DtpFromDate.Value, "/", "/") & "' and  IssuedDate<='" & FormatDMY_to_YMD(DtpTodate.Value, "/", "/") & "' "
    'sqlStr = sqlStr & " And CompanyID='" & companyID_ & "'"
    sqlStr = SelectWithCompanyID(sqlStr, True)
    Set colSearchResult_ = tmpS.BillingHeaderSearch(sqlStr)
    lvResult.ListItems.Clear
    For x = 1 To colSearchResult_.Count
        Set tmpBill = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpBill.ID & "'", tmpBill.BillingNo)
        tmpList.SubItems(1) = FormatYMD_to_DMY(tmpBill.IssuedDate, "/", "/")
        tmpList.SubItems(2) = Format(tmpBill.AmtB4VAT, "#,##0.00")
        tmpList.SubItems(3) = Format(tmpBill.PaidAmt, "#,##0.00")
        tmpList.SubItems(4) = Format(tmpBill.VATAmt, "#,##0.00")
        tmpList.SubItems(5) = Format(tmpBill.TotalBillAmt, "#,##0.00")
        tmpList.Checked = isExist(colOriSelProduct_, tmpList.Key) Or isExist(colItemAdd_, tmpList.Key)
        Set tmpList = Nothing
        Set tmpBill = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Label1(2).Caption = "ผลการค้นหา พบ " & lvResult.ListItems.Count & " รายการ"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub Form_Load()
    DtpFromDate.Value = TodayDate
    DtpTodate.Value = TodayDate
End Sub

Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
Dim selStr As String
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    selStr = Trim(curCustomer_.CustomerName)
    If Trim(txtCustomer.Text) = "" And Trim(selStr) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(selStr) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .ShowSubContact = True
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal ""
'                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_ Is Nothing Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set colItemAdd_ = Nothing
    Set colItemDel_ = Nothing
End Sub

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpList As ListItem
Dim CName As String
Set curBill_ = New BWGRepIvcBill
    If item.Checked = True Then
        colItemAdd_.Add colSearchResult_(item.Key), item.Key
        If isExist(colItemDel_, item.Key) Then
            colItemDel_.Remove item.Key
        End If
        If isExist(lvSelected.ListItems, item.Key) = False Then
            Set tmpList = lvSelected.ListItems.Add(, item.Key, item.Text)
            tmpList.SubItems(1) = item.SubItems(1)
            tmpList.SubItems(2) = item.SubItems(2)
            tmpList.SubItems(3) = item.SubItems(4)
            tmpList.SubItems(4) = item.SubItems(5)
            tmpList.SubItems(5) = tmpList
            With curBill_
            If Trim(curCustomer_.CustomerName) = "" Then
                MsgBox "ไม่พบข้อมูลลูกค้า กรุณารุบะรหัส หรือชื่อลูกค้าก่อน", vbExclamation
            Exit Sub
            End If
            CName = "บริษัท " & curCustomer_.CustomerName
                .BillingNo = tmpList
                .IssuedDate = item.SubItems(1)
                .AmtB4VAT = item.SubItems(2)
                .VATAmt = item.SubItems(4)
                .TotalBillAmt = item.SubItems(5)
                .CustomerName = CName
                .UserN = CurrentUser.LoginName
            End With
                curBill_.Save
        End If
    Else
            With curBill_
                .BillingNo = item.Text
                .UserN = CurrentUser.LoginName
            End With
        curBill_.Delete
        If isExist(colItemDel_, item.Key) = False Then
            colItemDel_.Add colSearchResult_(item.Key), item.Key
        End If
        If isExist(colItemAdd_, item.Key) Then
            colItemAdd_.Remove item.Key
        End If
        If isExist(lvSelected.ListItems, item.Key) = True Then
            lvSelected.ListItems.Remove item.Key
        End If
    End If
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub
Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = selCust
    If Not curCustomer_ Is Nothing Then
        txtCustomer.Text = curCustomer_.CustomerName
    Else
        txtCustomer.Text = ""
    End If
End Sub
'Public Sub setCallerForm(frm As Form, selCust As BWGCustomer, invID As String, CompanyID As String)
'Dim x%
'Dim tmpBill As BWGBillingHeader
'Dim tmpList As ListItem
'Dim tmpS As BWGSearchManager
'Dim criStr As String
'Dim strCond As String
'    Set callerFrm_ = frm
'    Set curCustomer_ = selCust
'    lbCustomer.Caption = selCust.CustomerName
'    invHeaderID_ = invID
'    companyID_ = CompanyID
'    Set colOriSelProduct_ = frm.getBillItems
'    lvSelected.ListItems.Clear
'
'    If Not colOriSelProduct_ Is Nothing Then
'        For x = 1 To colOriSelProduct_.Count
'
'            Set colSearchResult_ = Nothing
'            Set colSearchResult_ = New Collection
'            Set tmpS = New BWGSearchManager
'            strCond = "(BillingHeaderID Like '%" & colOriSelProduct_(x).ID & "%')"
'            Set colSearchResult_ = tmpS.BillingHeaderSearch(strCond)
'            Set tmpS = Nothing
'
'
'            Set tmpBill = colSearchResult_(1)
'            Set tmpList = lvSelected.ListItems.Add(, "'" & tmpBill.ID & "'", tmpBill.BillingNo)
'            tmpList.SubItems(1) = FormatYMD_to_DMY(tmpBill.IssuedDate, "/", "/")
'            tmpList.SubItems(2) = Format(tmpBill.TotalBillAmt, "#,##0.00")
'            tmpList.SubItems(3) = Format(tmpBill.PaidAmt, "#,##0.00")
'            Set tmpList = Nothing
'            Set tmpBill = Nothing
'        Next
'    End If
'    cmdSearch_Click
'    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
'End Sub

Public Sub AddSelectedProduct(selBill As BWGBillingHeader)
Dim tmpKey As String
Dim tmpList As ListItem
    tmpKey = Trim("'" & selBill.ID & "'")
    If isExist(colItemAdd_, tmpKey) Then Exit Sub
    colItemAdd_.Add selBill, tmpKey
    If isExist(colItemDel_, tmpKey) Then
        colItemDel_.Remove tmpKey
    End If
    If isExist(lvSelected.ListItems, tmpKey) = False Then
        Set tmpList = lvSelected.ListItems.Add(, tmpKey, selBill.BillingNo)
        tmpList.SubItems(1) = FormatYMD_to_DMY(selBill.IssuedDate, "/", "/")
        tmpList.SubItems(2) = Format(selBill.AmtB4VAT, "#,##0.00")
        tmpList.SubItems(3) = Format(selBill.PaidAmt, "#,##0.00")
        Set tmpList = Nothing
    End If
End Sub

Private Sub lvSelected_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim tmpKey As String

    With lvSelected
        If KeyCode = 46 And Not .SelectedItem Is Nothing Then
            selResult = MsgBox("ต้องการลบใบวางบิลนี้ออกหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = Trim(.SelectedItem.Key)
                curBill_.Delete
                If isExist(colItemDel_, tmpKey) = False Then
                    If isExist(colOriSelProduct_, tmpKey) Then
                        colItemDel_.Add colOriSelProduct_(tmpKey), tmpKey
                    ElseIf isExist(colItemAdd_, tmpKey) Then
                        colItemDel_.Add colItemAdd_(tmpKey), tmpKey
                        colItemAdd_.Remove tmpKey
                    Else
                        MsgBox "ไม่พบรายการใบวางบิลที่ต้องการลบ", vbExclamation
                    End If
                End If
                If isExist(lvSelected.ListItems, tmpKey) = True Then
                    lvSelected.ListItems.Remove tmpKey
                End If
            End If
        End If
    End With
End Sub

