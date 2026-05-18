VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWasteDataAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F149"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "อัปเดทรายการที่เลือก / ปิด"
      Height          =   375
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1290
      Width           =   2595
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   6990
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
      Top             =   1350
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2460
      Width           =   9705
      _ExtentX        =   17119
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Waste No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อของเสีย"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "EU Code"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "QuotationReq"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หมายเหตุ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "สถานะ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ปลายทาง"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1905
      Left            =   0
      TabIndex        =   5
      Top             =   5430
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   3360
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
         Text            =   "Waste No"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อของเสีย"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "EU Code"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Waste Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อของเสีย, Waste No"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   1710
      Width           =   1530
   End
   Begin VB.Label lbCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   990
      Width           =   570
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   9720
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
      TabIndex        =   9
      Top             =   5160
      Width           =   9735
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   9720
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
      TabIndex        =   8
      Top             =   2220
      Width           =   9735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   345
      TabIndex        =   7
      Top             =   1410
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกรายการกากของเสีย"
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
      TabIndex        =   6
      Top             =   90
      Width           =   2310
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
      Picture         =   "frWasteDataAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "frmWasteDataAdd"
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

Dim colMergeCollection As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim X%
Dim tmpKey As String
    If colOriSelProduct_ Is Nothing Then Set colOriSelProduct_ = New Collection
    For X = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(X).id & "'")
        If isExist(colOriSelProduct_, tmpKey) Then
            callerFrm_.RemoveWasteData colItemDel_(X).id
        End If
    Next
    For X = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(X).id & "'")
        If isExist(colOriSelProduct_, tmpKey) = False Then
            colOriSelProduct_.Add colItemAdd_(X), tmpKey
            callerFrm_.AddWasteDataItem colItemAdd_(X)
        End If
    Next
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpList As ListItem
Dim tmpWaste As BWGWasteDataCR
Dim tmpArr
Dim criStr1 As String, tmpStr$, curNameSpec As String
Dim sqlStr$, tSql2 As String, tSql3 As String
Dim X%, i As Integer
    If Trim(txtKey.Text) = "" Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtKey.SetFocus
        Exit Sub
    End If
    criStr1 = Replace(Trim(txtKey.Text), "*", "%")
    sqlStr = "(WasteNo Like '%" & criStr1 & "%' Or WasteName Like '%" & criStr1 & "%') "
    If Not curCustomer_ Is Nothing Then
        sqlStr = sqlStr & " And CustomerID='" & curCustomer_.id & "'"
    End If
                
    tSql3 = "SELECT WasteDataID FROM vw_WasteDataCRSearch "
    tSql3 = tSql3 & " WHERE " & sqlStr
    tSql2 = " AND (ISNULL(isQuotationReq, 'N') = 'Y') AND (Status <> 'J')"
    tSql2 = tSql2 & " AND (WasteDataID NOT IN (" & tSql3 & ")) "
    
    Set colSearchResult_ = tmpS.WasteDataCRSearch(sqlStr & " And Status IN ( 'A' ,'D', 'P', 'R')")
    lvResult.ListItems.Clear
    For X = 1 To colSearchResult_.Count
        Set tmpWaste = colSearchResult_(X)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpWaste.id & "'", tmpWaste.WasteNo)
        tmpList.SubItems(1) = tmpWaste.WasteName
        tmpList.SubItems(2) = tmpWaste.EUCode1 & "-" & tmpWaste.EUCode2 & "-" & tmpWaste.EUCode3
        tmpList.SubItems(3) = tmpWaste.WasteType
        If tmpWaste.isQuotationReq = False Then
            tmpList.SubItems(4) = ""
        Else
            tmpList.SubItems(4) = "ออกใบเสนอราคาเร่งด่วน"
        End If
        tmpList.SubItems(5) = tmpWaste.QuotationPrice
        Dim wState As String
        If Trim(tmpWaste.Status) = "A" Or Trim(tmpWaste.Status) = "IA" Then
            wState = "รับดำเนินการ"
        ElseIf Trim(tmpWaste.Status) = "P" Then
            wState = "ยังไม่สรุปผล"
        ElseIf Trim(tmpWaste.Status) = "R" Then
            wState = "รอตรวจรับ"
        ElseIf Trim(tmpWaste.Status) = "D" Then
            wState = "ยังไม่ส่งข้อมูล"
        ElseIf Trim(tmpWaste.Status) = "E" Then
             wState = "ยังไม่ส่งข้อมูล"
        Else
            wState = ""
        End If
         tmpList.SubItems(6) = wState
        tmpList.SubItems(7) = tmpWaste.CompanyCode
        tmpList.Checked = isExist(colOriSelProduct_, tmpList.Key) Or isExist(colItemAdd_, tmpList.Key)
        Set tmpList = Nothing
        Set tmpWaste = Nothing
    Next
    
    Set colMergeCollection = New Collection
    For i = 1 To colSearchResult_.Count
        colMergeCollection.Add colSearchResult_.Item(i)
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
        If colMergeCollection(Item.Index).Status <> "A" Then
            If CheckDiffType("", colMergeCollection(Item.Index).isQuotationReq) = True Then
                Item.Checked = False
                MsgBox "เวสต์ที่เลือกสถานะการขอออกใบเสนอราคาไม่เหมือนกันกรุณาเลือกเวสต์อื่น", vbExclamation, "Warning"
                Exit Sub
            Else
                colItemAdd_.Add colMergeCollection(Item.Index), Item.Key
            End If
        Else
            colItemAdd_.Add colMergeCollection(Item.Index), Item.Key
        End If
        If isExist(colItemDel_, Item.Key) Then
            colItemDel_.Remove Item.Key
        End If
        If isExist(lvSelected.ListItems, Item.Key) = False Then
            Set tmpList = lvSelected.ListItems.Add(, Item.Key, Item.Text)
            tmpList.SubItems(1) = Item.SubItems(1)
            tmpList.SubItems(2) = Item.SubItems(2)
        End If
    Else
        If isExist(colItemDel_, Item.Key) = False Then
            colItemDel_.Add colMergeCollection(Item.Index), Item.Key
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

Private Function CheckDiffType(ByVal wastedataID, wType As Boolean) As Boolean
    Dim i As Integer, tmpType As Boolean
    CheckDiffType = False: tmpType = False
    For i = 1 To colItemAdd_.Count
        tmpType = colItemAdd_.Item(i).isQuotationReq
        If wType <> tmpType Then
            CheckDiffType = True
            Exit For
        Else
            CheckDiffType = False
        End If
    Next i
End Function

Public Sub SetCallerForm(frm As Form, selCust As BWGCustomer)
Dim X%
Dim tmpWaste As BWGWasteDataCR
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim criStr As String
Dim strCond As String
    Set callerFrm_ = frm
    Set curCustomer_ = selCust
    lbCustomer.Caption = selCust.CustomerName
    Set colOriSelProduct_ = frm.getWastDataItems
    lvSelected.ListItems.Clear

    If Not colOriSelProduct_ Is Nothing Then
        For X = 1 To colOriSelProduct_.Count

            Set colSearchResult_ = Nothing
            Set colSearchResult_ = New Collection
            Set tmpS = New BWGSearchManager
            strCond = "(WasteDataID Like '%" & colOriSelProduct_(X).id & "%')"
            
            Set colSearchResult_ = tmpS.WasteDataCRSearch(strCond)
            Set tmpS = Nothing

            If colSearchResult_.Count > 0 Then
                Set tmpWaste = colSearchResult_(1)
                Set tmpList = lvSelected.ListItems.Add(, "'" & tmpWaste.id & "'", tmpWaste.WasteNo)
                tmpList.SubItems(1) = tmpWaste.WasteName
                tmpList.SubItems(2) = tmpWaste.EUCode1 & "-" & tmpWaste.EUCode2 & "-" & tmpWaste.EUCode3
                tmpList.SubItems(3) = tmpWaste.WasteType
                Set tmpList = Nothing
                Set tmpWaste = Nothing
            End If
        Next
    End If
    cmdSearch_Click
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub AddSelectedProduct(selWaste As BWGWasteDataCR)
Dim tmpKey As String
Dim tmpList As ListItem
    tmpKey = Trim("'" & selWaste.id & "'")
    If isExist(colItemAdd_, tmpKey) Then Exit Sub
    colItemAdd_.Add selWaste, tmpKey
    If isExist(colItemDel_, tmpKey) Then
        colItemDel_.Remove tmpKey
    End If
    If isExist(lvSelected.ListItems, tmpKey) = False Then
        Set tmpList = lvSelected.ListItems.Add(, tmpKey, selWaste.WasteNo)
        tmpList.SubItems(1) = selWaste.WasteName
        tmpList.SubItems(2) = selWaste.EUCode1 & "-" & selWaste.EUCode2 & "-" & selWaste.EUCode3
        tmpList.SubItems(3) = selWaste.WasteType
        Set tmpList = Nothing
    End If
End Sub

Private Sub lvSelected_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim tmpKey As String
    With lvSelected
        If KeyCode = 46 And Not .SelectedItem Is Nothing Then
            selResult = MsgBox("ต้องการลบกากของเสียนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
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
