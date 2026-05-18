VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmREIVAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "ไม่เลือก"
      Height          =   315
      Left            =   1230
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เลือกทั้งหมด"
      Height          =   315
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1245
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   1590
      Width           =   3855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   930
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "อัปเดทรายการที่เลือก / ปิด"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1350
      Width           =   2415
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1770
      Width           =   2415
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "*"
      Top             =   1980
      Width           =   3855
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   2910
      Width           =   10335
      _ExtentX        =   18230
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
         Text            =   "วันที่ Invoice"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ Invoice"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่ Manifest"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ยอดรวมก่อน VAT"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "VAT"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "ยอดรวมทั้งสิ้น"
         Object.Width           =   2999
      EndProperty
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   5850
      Width           =   10335
      _ExtentX        =   18230
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
         Text            =   "วันที่ขน"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ Manifest"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รายการของเสีย"
         Object.Width           =   5556
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รหัส"
         Object.Width           =   2328
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ทะเบียนรถ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "น้ำหนัก"
         Object.Width           =   1826
      EndProperty
   End
   Begin Project1.ctlDate dtFrom 
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin Project1.ctlDate dtTo 
      Height          =   345
      Left            =   4830
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   2130
      TabIndex        =   20
      Top             =   1650
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึง"
      Height          =   195
      Index           =   3
      Left            =   4590
      TabIndex        =   19
      Top             =   1260
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ช่วงวันที่ขน"
      Height          =   195
      Index           =   2
      Left            =   1890
      TabIndex        =   18
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อของเสีย, Waste No, เลขที่ Manifest, ทะเบียนรถ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   17
      Top             =   2310
      Width           =   3495
   End
   Begin VB.Label lbCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2790
      TabIndex        =   16
      Top             =   930
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อบริษัท"
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   930
      Width           =   615
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   12000
      Y1              =   5580
      Y2              =   5580
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
      TabIndex        =   14
      Top             =   5610
      Width           =   10365
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   12000
      Y1              =   2640
      Y2              =   2640
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
      TabIndex        =   13
      Top             =   2670
      Width           =   10365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   1305
      TabIndex        =   12
      Top             =   2040
      Width           =   1365
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
      TabIndex        =   11
      Top             =   90
      Width           =   1380
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frREIVAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "frmREIVAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim curCompany_ As BWGCompanyConfig
Dim curCustomer_ As BWGCustomer
Dim colOriSelProduct_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim refID_ As String
Dim showAlert_ As Boolean
Dim callerFrm_ As Form

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim x%
Dim tmpKey As String
    If colOriSelProduct_ Is Nothing Then Set colOriSelProduct_ = New Collection
    For x = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(x).ID & "_" & colItemDel_(x).MenifestID & "'")
        If isExist(colOriSelProduct_, tmpKey) Then
            callerFrm_.RemovePrintDetail colItemDel_(x).ID, colItemDel_(x).MenifestID
        End If
    Next
    For x = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(x).ID & "'")
        If isExist(colOriSelProduct_, tmpKey) = False Then
            colOriSelProduct_.Add colItemAdd_(x), tmpKey
            callerFrm_.AddPrintDetail colItemAdd_(x)
        End If
    Next
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim tmpS As New BWGSearchManager
Dim tmpList As ListItem
Dim tmpInvDet As BWGInvData
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
    sqlStr = "(BillingNo Like '%" & criStr1 & "%' Or DocNo Like '%" & criStr1 & "%' And InvDate >= '" & dtFrom.ValueYMD & "' And InvDate <= '" & dtTo.ValueYMD & "')"
    If Not curCompany_ Is Nothing Then
        sqlStr = sqlStr & " And CompanyID='" & curCompany_.ID & "'"
    End If
    If Not curCustomer_ Is Nothing Then
        sqlStr = sqlStr & " And CustomerID='" & curCustomer_.ID & "'"
    End If
    Set colSearchResult_ = tmpS.InvSearch(sqlStr)
    lvResult.ListItems.Clear
    For x = 1 To colSearchResult_.Count
        Set tmpInvDet = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpInvDet.InvID & "_" & tmpInvDet.MenifestID & "'", FormatYMD_to_DMY(tmpInvDet.InvDate))
        tmpList.SubItems(1) = tmpInvDet.BillingNo
        tmpList.SubItems(2) = tmpInvDet.DocNo
        tmpList.SubItems(3) = Format(tmpInvDet.ManiAmtB4VAT, "0.00")
        tmpList.SubItems(4) = Format(tmpInvDet.MeniVATAmt, "0.00")
        tmpList.SubItems(5) = Format(tmpInvDet.TotalAmtToPay, "0.00")
        tmpList.Checked = isExist(colOriSelProduct_, tmpList.Key) Or isExist(colItemAdd_, tmpList.Key)
        Set tmpList = Nothing
        Set tmpInvDet = Nothing
    Next
    If lvResult.ListItems.Count = 0 And showAlert_ Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Label1(2).Caption = "ผลการค้นหา พบ " & lvResult.ListItems.Count & " รายการ"
End Sub

Private Sub Command1_Click()
Dim x%
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked = False Then
            lvResult.ListItems(x).Checked = True
            lvResult_ItemCheck lvResult.ListItems(x)
        End If
    Next
End Sub

Private Sub Command2_Click()
Dim x%
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked = True Then
            lvResult.ListItems(x).Checked = False
            lvResult_ItemCheck lvResult.ListItems(x)
        End If
    Next
End Sub

Private Sub dtFrom_GotFocus()
    showAlert_ = True
End Sub

Private Sub dtTo_GotFocus()
    showAlert_ = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.ActiveControl.Name <> "txtCustomer" Then
            cmdSearch_Click
        Else
            SendKeys "{Tab}"
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
    If item.Checked = True Then
        colItemAdd_.Add colSearchResult_(item.Key), item.Key
        If isExist(colItemDel_, item.Key) Then
            colItemDel_.Remove item.Key
        End If
        If isExist(lvSelected.ListItems, item.Key) = False Then
            Set tmpList = lvSelected.ListItems.Add(, item.Key, item.Text)
            tmpList.SubItems(1) = item.SubItems(1)
            tmpList.SubItems(2) = item.SubItems(2)
            tmpList.SubItems(3) = item.SubItems(3)
            tmpList.SubItems(4) = item.SubItems(4)
            tmpList.SubItems(5) = item.SubItems(5)
        End If
    Else
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
    Label1(1).Caption = "รายการที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub setCallerForm(frm As Form, selCompany As BWGCompanyConfig)
Dim x%
Dim tmpInvDet As BWGInvData
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim criStr As String
Dim strCond As String
    Set callerFrm_ = frm
    Set curCompany_ = selCompany
    lbCustomer.Caption = selCompany.CompanyNameTH
    Set colOriSelProduct_ = frm.GetJobDetails
    refID_ = frm.RefID
    lvSelected.ListItems.Clear

    If Not colOriSelProduct_ Is Nothing Then
        For x = 1 To colOriSelProduct_.Count

            Set colSearchResult_ = Nothing
            Set colSearchResult_ = New Collection
            Set tmpInvDet = colOriSelProduct_(x)
            Set tmpList = lvSelected.ListItems.Add(, "'" & tmpInvDet.InvID & "_" & tmpInvDet.MenifestID & "'", FormatYMD_to_DMY(tmpInvDet.InvDate))
            tmpList.SubItems(1) = tmpInvDet.BillingNo
            tmpList.SubItems(2) = tmpInvDet.DocNo
            tmpList.SubItems(3) = Format(tmpInvDet.ManiAmtB4VAT, "0.00")
            tmpList.SubItems(4) = Format(tmpInvDet.MeniVATAmt, "0.00")
            tmpList.SubItems(5) = Format(tmpInvDet.TotalAmtToPay, "0.00")
            Set tmpList = Nothing
            Set tmpInvDet = Nothing
        Next
    End If
    showAlert_ = False
    cmdSearch_Click
    Label1(1).Caption = "รายการที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

'Public Sub AddSelectedProduct(selWaste As BWGWasteDataCR)
'Dim tmpKey As String
'Dim tmpList As ListItem
'    tmpKey = Trim("'" & selWaste.ID & "'")
'    If isExist(colItemAdd_, tmpKey) Then Exit Sub
'    colItemAdd_.Add selWaste, tmpKey
'    If isExist(colItemDel_, tmpKey) Then
'        colItemDel_.Remove tmpKey
'    End If
'    If isExist(lvSelected.ListItems, tmpKey) = False Then
'        Set tmpList = lvSelected.ListItems.Add(, tmpKey, selWaste.WasteNo)
'        tmpList.SubItems(1) = selWaste.WasteName
'        tmpList.SubItems(2) = selWaste.EUCode1 & "-" & selWaste.EUCode2 & "-" & selWaste.EUCode3
'        Set tmpList = Nothing
'    End If
'End Sub

Private Sub lvSelected_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim tmpKey As String
    With lvSelected
        If KeyCode = 46 And Not .SelectedItem Is Nothing Then
            selResult = MsgBox("ต้องการลบกากของเสียนี้ออกจากใบแจ้งหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = Trim(.SelectedItem.Key)
                If isExist(colItemDel_, tmpKey) = False Then
                    If isExist(colOriSelProduct_, tmpKey) Then
                        colItemDel_.Add colOriSelProduct_(tmpKey), tmpKey
                    ElseIf isExist(colItemAdd_, tmpKey) Then
                        colItemDel_.Add colItemAdd_(tmpKey), tmpKey
                        colItemAdd_.Remove tmpKey
                    Else
                        MsgBox "ไม่พบรายการสินค้าที่ต้องการลบ", vbExclamation
                    End If
                End If
                If isExist(lvSelected.ListItems, tmpKey) = True Then
                    lvSelected.ListItems.Remove tmpKey
                End If
            End If
        End If
    End With
End Sub

Private Sub txtKey_GotFocus()
    showAlert_ = True
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
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

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub

