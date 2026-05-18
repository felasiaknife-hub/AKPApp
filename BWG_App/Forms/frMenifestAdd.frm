VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifestAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F057"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   12
      Top             =   2460
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   4683
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "อัปเดทรายการที่เลือก / ปิด"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1290
      Width           =   2595
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   2595
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      Text            =   "*"
      Top             =   1350
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   10755
      _ExtentX        =   18971
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Menifest No"
         Object.Width           =   3545
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ขน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "สถานที่ขน"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "ยอดเงิน"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบ Menifest, ทะเบียนรถ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2670
      TabIndex        =   11
      Top             =   1710
      Width           =   2070
   End
   Begin VB.Label lbCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2700
      TabIndex        =   10
      Top             =   990
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกหนี้"
      Height          =   195
      Index           =   0
      Left            =   1950
      TabIndex        =   9
      Top             =   990
      Width           =   615
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   11000
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
      Width           =   10785
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   11000
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
      Width           =   10785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   1215
      TabIndex        =   6
      Top             =   1410
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาใบ Menifest ที่ต้องการวางบิล"
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
      Width           =   3420
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
      Picture         =   "frMenifestAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "frmMenifestAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim curBillHeader_ As BWGBillingHeader
Dim colOriSelProduct_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim callerFrm_ As Form
Dim companyID_ As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim x%
Dim tmpKey As String
    If colOriSelProduct_ Is Nothing Then Set colOriSelProduct_ = New Collection
    For x = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(x).ID & "'")
        If isExist(colOriSelProduct_, tmpKey) Then
            callerFrm_.RemoveMenifest colItemDel_(x).ID
        End If
    Next
    For x = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(x).ID & "'")
        If isExist(colOriSelProduct_, tmpKey) = False Then
            colOriSelProduct_.Add colItemAdd_(x), tmpKey
            callerFrm_.AddMenifest colItemAdd_(x)
        End If
    Next
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim tmpS As New BWGSearchManager
Dim tmpList As ListItem, curTimeTableNo$
Dim tmpMenifest As BWGMenifest
Dim tmpArr, tmpSite As BWGCustomerSite
Dim criStr1 As String, tmpStr$, curNameSpec As String
Dim sqlStr$, tmpNode As Node, tmpChild As Node
Dim sqlStr2$
Dim x%, checkCount%, y%
Dim menifestID As String
    If Trim(txtKey.Text) = "" Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtKey.SetFocus
        Exit Sub
    End If
    criStr1 = Replace(Trim(txtKey.Text), "*", "%")
    sqlStr = "(DocNo Like '%" & criStr1 & "%' Or MenifestNo Like '%" & criStr1 & "%' Or CarRegisID Like '%" & criStr1 & "%' ) And (isCanceled<>'Y' Or isCanceled is null) And (BillingHeaderID='" & curBillHeader_.ID & "' Or BillingHeaderID is null) "
    If Not curCustomer_ Is Nothing Then
        sqlStr = sqlStr & " And (CustomerID='" & curCustomer_.ID & "' Or ParentCustID='" & curCustomer_.ID & "')"
    End If
    sqlStr = sqlStr & " AND (isInv <> 'Y') "
    'If Not curBillHeader_ Is Nothing Then
    '    sqlStr = sqlStr & " And isClosed='Y'"
    'End If
    menifestID = ""
        For x = 1 To colOriSelProduct_.Count
        If menifestID = "" Then
            menifestID = "'" & colOriSelProduct_(x).ID & "'"
        Else
            menifestID = menifestID & ",'" & colOriSelProduct_(x).ID & "'"
        End If
    Next
    sqlStr = sqlStr & " AND CompanyID IN ('" & companyID_ & "')"
    If Trim(menifestID) <> "" Then sqlStr2 = " MenifestID IN (" & menifestID & ")"
'    sqlStr = SelectWithCompanyID(sqlStr, True)
    If Trim(menifestID) <> "" Then
        Set colSearchResult_ = tmpS.MenifestSearchUnion(sqlStr, sqlStr2)
    Else
        Set colSearchResult_ = tmpS.MenifestSearch(sqlStr)
    End If
    tvResult.Nodes.Clear
    For x = 1 To colSearchResult_.Count
        Set tmpMenifest = colSearchResult_(x)
        If Trim(tmpMenifest.TimeTableNo) <> "" Then
            If Trim(curTimeTableNo) <> Trim(tmpMenifest.TimeTableNo) Then
                If isExist(tvResult.Nodes, "'" & tmpMenifest.TimeTableNo & "'") Then
                    Set tmpNode = tvResult.Nodes("'" & tmpMenifest.TimeTableNo & "'")
                Else
                    Set tmpNode = tvResult.Nodes.Add(, , "'" & tmpMenifest.TimeTableNo & "'", "ใบคุม : " & tmpMenifest.TimeTableNo & ", ทะเบียนรถ " & tmpMenifest.CarRegisID & ", ค่าขนส่ง : " & Format(tmpMenifest.TripTranFee + tmpMenifest.TripEquipFee + tmpMenifest.TripMoreWorkerFee + tmpMenifest.TripOtherFee, "#,##0.00") & " บาท")
                End If
                curTimeTableNo = Trim(tmpMenifest.TimeTableNo)
            Else
                Set tmpNode = tvResult.Nodes("'" & tmpMenifest.TimeTableNo & "'")
            End If
            Set tmpChild = tvResult.Nodes.Add(tmpNode, tvwChild, "'" & tmpMenifest.ID & "'", "เลขที่ : " & tmpMenifest.DocNo & " , Manifest No : " & tmpMenifest.MenifestNo & " , วันที่ขน : " & FormatYMD_to_DMY(tmpMenifest.WorkDate, "/", "/") & " , ยอดเงิน : " & Format(tmpMenifest.InCollectFee + tmpMenifest.InEquipFee + tmpMenifest.InOtherFee + tmpMenifest.IntransFee + tmpMenifest.InWorkerFee, "#,##0.00"))
            If tmpMenifest.isClosed = False Then
                tmpChild.ForeColor = vbRed
            End If
            tmpNode.Expanded = True
            If isExist(colOriSelProduct_, tmpChild.Key) Then
                tmpChild.Checked = True
                Set tmpNode = tmpChild.FirstSibling
                If tmpNode.Checked Then checkCount = 1
                For y = 1 To tmpChild.Parent.Children - 1
                    If tmpNode.Next.Checked = True Then
                        checkCount = checkCount + 1
                    End If
                Next
                If checkCount = tmpChild.Parent.Children Then tmpChild.Parent.Checked = True
            End If
            
            Set tmpNode = Nothing
        End If
        Set tmpMenifest = Nothing
        Set tmpChild = Nothing
    Next
    If tvResult.Nodes.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    Else
        tvResult.Nodes(1).Selected = True
    End If
    Label1(2).Caption = "ผลการค้นหา พบ " & tvResult.Nodes.Count & " รายการ"
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

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'MsgBox ColumnHeader.Width
End Sub

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpList As ListItem
Dim x%, selWasteID$
    If item.Checked = True Then
        selWasteID = Trim(item.Key)
        For x = 1 To lvSelected.ListItems.Count
            If Trim(selWasteID) = Trim(lvSelected.ListItems(x).Key) Then
                MsgBox "Menifest รายการนี้ถูกเลือกแล้วไม่สามารถเลือกซ้ำได้", vbExclamation
                item.Checked = False
                Exit Sub
            End If
        Next
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
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub SetCallerForm(frm As Form, selCust As BWGCustomer, curBillHeader As BWGBillingHeader, CompanyID As String)
Dim x%, tmpSite As BWGCustomerSite
Dim tmpMenifest As BWGMenifest
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim criStr As String
Dim strCond As String
    Set callerFrm_ = frm
    Set curCustomer_ = selCust
    lbCustomer.Caption = selCust.CustomerName
    Set colOriSelProduct_ = frm.GetMenifests
    lvSelected.ListItems.Clear
    Set curBillHeader_ = curBillHeader
    companyID_ = "776EDA5830F744A6B3551348D851BC22"
    If Not colOriSelProduct_ Is Nothing Then
        For x = 1 To colOriSelProduct_.Count
            Set tmpMenifest = colOriSelProduct_(x)
            Set tmpList = lvSelected.ListItems.Add(, "'" & tmpMenifest.ID & "'", tmpMenifest.DocNo)
            tmpList.SubItems(1) = tmpMenifest.MenifestNo
            tmpList.SubItems(2) = FormatYMD_to_DMY(tmpMenifest.WorkDate, "/", "/")
            If isExist(curCustomer_.SiteAddresses, "'" & tmpMenifest.CustSiteID & "'") Then
                Set tmpSite = curCustomer_.SiteAddresses("'" & tmpMenifest.CustSiteID & "'")
                tmpList.SubItems(3) = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName
                Set tmpSite = Nothing
            Else
                tmpList.SubItems(3) = " "
            End If
            tmpList.SubItems(4) = tmpMenifest.InCollectFee + tmpMenifest.InEquipFee + tmpMenifest.InOtherFee + tmpMenifest.IntransFee + tmpMenifest.InWorkerFee
            Set tmpList = Nothing
            Set tmpMenifest = Nothing
        Next
    End If
    cmdSearch_Click
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub AddSelectedProduct(selMenifest As BWGMenifest)
Dim tmpKey As String, tmpSite As BWGCustomerSite
Dim tmpList As ListItem
    tmpKey = Trim("'" & selMenifest.ID & "'")
    If isExist(colItemAdd_, tmpKey) Then Exit Sub
    colItemAdd_.Add selMenifest, tmpKey
    If isExist(colItemDel_, tmpKey) Then
        colItemDel_.Remove tmpKey
    End If
    'If isExist(lvSelected.ListItems, tmpKey) = False Then
        'Set tmpList = lvResult.ListItems.Add(, "'" & selMenifest.ID & "'", selMenifest.DocNo)
        'tmpList.SubItems(1) = selMenifest.MenifestNo
        'tmpList.SubItems(2) = FormatYMD_to_DMY(selMenifest.workDate, "/", "/")
        'If isExist(curCustomer_.SiteAddresses, "'" & selMenifest.CustSiteID & "'") Then
        '    Set tmpSite = curCustomer_.SiteAddresses("'" & selMenifest.CustSiteID & "'")
        '    tmpList.SubItems(3) = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName
        '    Set tmpSite = Nothing
        'Else
        '    tmpList.SubItems(3) = " "
        'End If
        'tmpList.SubItems(4) = selMenifest.InCollectFee + selMenifest.InEquipFee + selMenifest.InOtherFee + selMenifest.IntransFee + selMenifest.InWorkerFee
        'Set tmpList = Nothing
    'End If
End Sub

Private Sub lvSelected_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim tmpKey As String
    With lvSelected
        If KeyCode = 46 And Not .SelectedItem Is Nothing Then
            selResult = MsgBox("ต้องการ Menifest นี้ออกจากใบแจ้งหนี้หรือไม่?", vbQuestion + vbYesNo)
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

Private Sub tvResult_DblClick()
On Error GoTo ErrD
Dim tmpMID As String, tmpCol As Collection, tmpS As New BWGSearchManager
Dim id1$, id2$, CarType$, aumphurID1$, aumphurID2$, truckTypeID1$
Dim selTxt As String
    If tvResult.SelectedItem.Parent Is Nothing Then
        tmpMID = Replace(tvResult.SelectedItem.Child.Key, "'", "")
    Else
        tmpMID = Replace(tvResult.SelectedItem.Key, "'", "")
    End If
    Set tmpCol = tmpS.MenifestSearch("MenifestID='" & tmpMID & "'")
    With tmpCol(1)
        'id1 = .JobDataCarID
        'id2 = .tripNo
        id1 = .TimeTableNo
        CarType = .TruckSubTypeName & ", " & .TruckTypeName
        aumphurID1 = Trim(.CustSiteAumphurID)
        aumphurID2 = Trim(.TSDFAumphurID)
        truckTypeID1 = Trim(.TruckTypeID)
    End With
    'Set tmpCol = tmpS.JobDataTimeTableSearch("JobDataCarID='" & id1 & "' And TripNo=" & id2)
    Set tmpCol = tmpS.JobDataTimeTableSearch("TimeTableNo='" & id1 & "' And isMain='Y' And isCancel ='N'")
    If tmpCol.Count > 0 Then
        Load frmTripControl
        With frmTripControl
            If InStr(1, tvResult.SelectedItem.Text, "ทะเบียนรถ") > 0 Then
                selTxt = tvResult.SelectedItem.Text
            Else
                selTxt = tvResult.SelectedItem.Parent.Text
            End If
            .SetTripControl tmpCol(1), lbCustomer, CarType, Trim(Replace(Mid(selTxt, InStr(1, selTxt, "ทะเบียนรถ"), InStr(InStr(1, selTxt, "ทะเบียนรถ"), selTxt, ",") - InStr(1, selTxt, "ทะเบียนรถ")), "ทะเบียนรถ", "")), aumphurID1, aumphurID2, truckTypeID1
            .Show 1
            tvResult.SelectedItem.Expanded = True
        End With
    End If
    Exit Sub
ErrD:
    MsgBox err.Description, vbCritical, "Error"
    Exit Sub
End Sub

Private Sub tvResult_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim x%, tmpNode As Node, checkCount%
Dim selWasteID$
    If Node.Parent Is Nothing Then
        For x = 1 To tvResult.Nodes.Count
            Set tmpNode = tvResult.Nodes(x)
            If Not tmpNode.Parent Is Nothing Then
                If Trim(tmpNode.Parent.Key) = Trim(Node.Key) And tmpNode.ForeColor <> vbRed Then
                    tmpNode.Checked = Node.Checked
                    checkCount = checkCount + 1
                End If
            End If
        Next
        If checkCount = 0 And Node.Checked = True Then
            MsgBox "ท่านไม่สามารถเลือก Manifest ในรถเที่ยวนี้ได้เพราะยังไม่มีรายการใดถูกปิดงาน", vbExclamation
            Node.Selected = True
            SendKeys " "
            Exit Sub
        End If
    Else
        If Node.Checked = False Then
            Node.Parent.Checked = False
        Else
            If Node.ForeColor = vbRed Then
                MsgBox "ท่านไม่สามารถเลือกใบ Manifest นี้ได้เพราะยังไม่ถูกปิดงาน", vbExclamation
                Node.Selected = True
                SendKeys " "
                Exit Sub
            Else
                Set tmpNode = Node.FirstSibling
                If tmpNode.Checked Then checkCount = 1
                For x = 1 To Node.Parent.Children - 1
                    If tmpNode.Next.Checked = True Then
                        checkCount = checkCount + 1
                    End If
                Next
                If checkCount = Node.Parent.Children Then Node.Parent.Checked = True
            End If
        End If
    End If
    AddSelManifest Node.Key, Node.Checked
End Sub

Private Sub AddSelManifest(selKey As String, isChecked As Boolean)
Dim tmpList As ListItem, isSingle As Boolean, tmpMf As BWGMenifest
Dim x%, selWasteID$, tmpCol As Collection, tmpKey As String
Dim tmpSite As BWGCustomerSite
    Set tmpCol = New Collection
    If isExist(colSearchResult_, selKey) Then
        tmpCol.Add colSearchResult_(selKey), selKey
        isSingle = True
    Else
        For x = 1 To colSearchResult_.Count
            If "'" & Trim(colSearchResult_(x).TimeTableNo) & "'" = selKey And colSearchResult_(x).isClosed Then
                tmpCol.Add colSearchResult_(x), "'" & colSearchResult_(x).ID & "'"
            End If
        Next
        isSingle = False
    End If
    
    If isChecked = True Then
        If isSingle Then
            For x = 1 To lvSelected.ListItems.Count
                If isExist(tmpCol, Trim(lvSelected.ListItems(x).Key)) Then
                    MsgBox "Menifest รายการนี้ถูกเลือกแล้วไม่สามารถเลือกซ้ำได้", vbExclamation
                    SendKeys " "
                    Exit Sub
                End If
            Next
        End If
        For x = 1 To tmpCol.Count
            tmpKey = "'" & tmpCol(x).ID & "'"
            If isExist(colItemAdd_, tmpKey) = False Then
                Set tmpMf = tmpCol(x)
                colItemAdd_.Add tmpCol(x), tmpKey
                If isExist(colItemDel_, tmpKey) Then
                    colItemDel_.Remove tmpKey
                End If
                If isExist(lvSelected.ListItems, tmpKey) = False Then
                    Set tmpList = lvSelected.ListItems.Add(, tmpKey, tmpMf.DocNo)
                    tmpList.SubItems(1) = tmpMf.MenifestNo
                    tmpList.SubItems(2) = FormatYMD_to_DMY(tmpMf.WorkDate, "/", "/")
                    If isExist(curCustomer_.SiteAddresses, "'" & tmpMf.CustSiteID & "'") Then
                        Set tmpSite = curCustomer_.SiteAddresses("'" & tmpMf.CustSiteID & "'")
                        tmpList.SubItems(3) = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName
                        Set tmpSite = Nothing
                    Else
                        tmpList.SubItems(3) = " "
                    End If
                    tmpList.SubItems(4) = tmpMf.InCollectFee + tmpMf.InEquipFee + tmpMf.InOtherFee + tmpMf.IntransFee + tmpMf.InWorkerFee
                End If
            End If
        Next
    Else
        For x = 1 To tmpCol.Count
            tmpKey = "'" & tmpCol(x).ID & "'"
            If isExist(colItemDel_, tmpKey) = False Then
                colItemDel_.Add tmpCol(tmpKey), tmpKey
            End If
            If isExist(colItemAdd_, tmpKey) Then
                colItemAdd_.Remove tmpKey
            End If
            If isExist(lvSelected.ListItems, tmpKey) = True Then
                lvSelected.ListItems.Remove tmpKey
            End If
        Next
    End If
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
    Set tmpCol = Nothing
End Sub
