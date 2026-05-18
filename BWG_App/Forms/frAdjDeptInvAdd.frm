VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAdjDeptInvAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F004"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "อัปเดทรายการที่เลือก / ปิด"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1290
      Width           =   2595
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   2595
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "*"
      Top             =   1290
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2460
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
      NumItems        =   4
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
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   12
      Top             =   5400
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
      NumItems        =   4
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
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดเงินที่แสดงในหน้าจอนี้เป็นยอดรวม VAT และหักภาษี ณ. ที่จ่ายแล้ว"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   13
      Top             =   1950
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบวางบิล"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label lbCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   10
      Top             =   990
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   9
      Top             =   990
      Width           =   570
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   8700
      Y1              =   5130
      Y2              =   5130
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
      TabIndex        =   8
      Top             =   5160
      Width           =   8655
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   8700
      Y1              =   2190
      Y2              =   2190
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
      TabIndex        =   7
      Top             =   2220
      Width           =   8655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   345
      TabIndex        =   6
      Top             =   1350
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
      TabIndex        =   5
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
      Picture         =   "frAdjDeptInvAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "frmAdjDeptInvAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim colOriSelProduct_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim callerFrm_ As Form
Dim invHeaderID_ As String
Dim companyID_ As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim x%
Dim tmpKey As String
    If colOriSelProduct_ Is Nothing Then Set colOriSelProduct_ = New Collection
    For x = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(x).ID & "'")
        If isExist(colOriSelProduct_, tmpKey) Then
            callerFrm_.RemoveBillData colItemDel_(x).ID
        End If
    Next
    For x = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(x).ID & "'")
        If isExist(colOriSelProduct_, tmpKey) = False Then
            colOriSelProduct_.Add colItemAdd_(x), tmpKey
            callerFrm_.AddBillItem colItemAdd_(x)
        End If
    Next
    Unload Me
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set colItemAdd_ = Nothing
    Set colItemDel_ = Nothing
End Sub

Private Sub lvResult_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim tmpList As ListItem
    If Item.Checked = True Then
        colItemAdd_.Add colSearchResult_(Item.Key), Item.Key
        If isExist(colItemDel_, Item.Key) Then
            colItemDel_.Remove Item.Key
        End If
        If isExist(lvSelected.ListItems, Item.Key) = False Then
            Set tmpList = lvSelected.ListItems.Add(, Item.Key, Item.Text)
            tmpList.SubItems(1) = Item.SubItems(1)
            tmpList.SubItems(2) = Item.SubItems(2)
            tmpList.SubItems(3) = Item.SubItems(3)
        End If
    Else
        If isExist(colItemDel_, Item.Key) = False Then
            colItemDel_.Add colSearchResult_(Item.Key), Item.Key
        End If
        If isExist(colItemAdd_, Item.Key) Then
            colItemAdd_.Remove Item.Key
        End If
        If isExist(lvSelected.ListItems, Item.Key) = True Then
            lvSelected.ListItems.Remove Item.Key
        End If
    End If
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub SetCallerForm(frm As Form, selCust As BWGCustomer, invID As String, companyID As String)
Dim x%
Dim tmpBill As BWGBillingHeader
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim criStr As String
Dim strCond As String
    Set callerFrm_ = frm
    Set curCustomer_ = selCust
    lbCustomer.Caption = selCust.CustomerName
    invHeaderID_ = invID
    companyID_ = companyID
    Set colOriSelProduct_ = frm.getBillItems
    lvSelected.ListItems.Clear

    If Not colOriSelProduct_ Is Nothing Then
        For x = 1 To colOriSelProduct_.Count

            Set colSearchResult_ = Nothing
            Set colSearchResult_ = New Collection
            Set tmpS = New BWGSearchManager
            strCond = "(BillingHeaderID Like '%" & colOriSelProduct_(x).ID & "%')"
            Set colSearchResult_ = tmpS.BillingHeaderSearch(strCond)
            Set tmpS = Nothing


            Set tmpBill = colSearchResult_(1)
            Set tmpList = lvSelected.ListItems.Add(, "'" & tmpBill.ID & "'", tmpBill.BillingNo)
            tmpList.SubItems(1) = FormatYMD_to_DMY(tmpBill.IssuedDate, "/", "/")
            tmpList.SubItems(2) = Format(tmpBill.TotalBillAmt, "#,##0.00")
            tmpList.SubItems(3) = Format(tmpBill.PaidAmt, "#,##0.00")
            Set tmpList = Nothing
            Set tmpBill = Nothing
        Next
    End If
    cmdSearch_Click
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

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
            selResult = MsgBox("ต้องการลบใบวางบิลนี้ออกจากการพิมพ์ใบกำกับ/ใบเสร็จ หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = Trim(.SelectedItem.Key)
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
