VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTripControlAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F132"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2670
      TabIndex        =   14
      Text            =   "*"
      Top             =   1620
      Width           =   4275
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   900
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "อัปเดทรายการที่เลือก / ปิด"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   2595
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1740
      Width           =   2595
   End
   Begin WasteManagment.ctlDate dtFrom 
      Height          =   345
      Left            =   2670
      TabIndex        =   0
      Top             =   1230
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate dtTo 
      Height          =   345
      Left            =   5160
      TabIndex        =   1
      Top             =   1230
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   12015
      _ExtentX        =   21193
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "บริษัทรับกำจัด"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ใบคุม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ขน"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ประเภทรถ"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เลขทะเบียน"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "คนขับ"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ต้นทาง"
         Object.Width           =   2355
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ปลายทาง"
         Object.Width           =   2355
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "ค่าบริการ"
         Object.Width           =   2143
      EndProperty
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   12015
      _ExtentX        =   21193
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "บริษัทรับกำจัด"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ใบคุม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ขน"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ประเภทรถ"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เลขทะเบียน"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "คนขับ"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ต้นทาง"
         Object.Width           =   2355
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ปลายทาง"
         Object.Width           =   2355
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "ค่าบริการ"
         Object.Width           =   2143
      EndProperty
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   0
      X2              =   13275
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   3
      Left            =   1170
      TabIndex        =   16
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขทะเบียนรถ, ชื่อคนขับ, เลขที่ใบคุม"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2670
      TabIndex        =   15
      Top             =   1950
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   13
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lbCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2700
      TabIndex        =   12
      Top             =   930
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ Vendor"
      Height          =   195
      Index           =   0
      Left            =   1770
      TabIndex        =   11
      Top             =   930
      Width           =   780
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
      TabIndex        =   10
      Top             =   5160
      Width           =   12075
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   13275
      Y1              =   2160
      Y2              =   2160
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
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   12075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ช่วงวันที่ขนตั้งแต่"
      Height          =   195
      Index           =   7
      Left            =   1410
      TabIndex        =   8
      Top             =   1290
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาใบคุมที่ต้องการทำจ่าย Vendor"
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
      TabIndex        =   7
      Top             =   90
      Width           =   3540
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
      Left            =   -60
      Picture         =   "frTripControlAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "frmTripControlAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curVendor_ As BWGTransportVendor
Dim curBillHeader_ As BWGBillPayHeader
Dim colOriSelProduct_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim callerFrm_ As Form

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim X%
Dim tmpKey As String
    If colOriSelProduct_ Is Nothing Then Set colOriSelProduct_ = New Collection
    For X = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(X).TimeTableNo & "_" & colItemDel_(X).TripNo & "'")
        If isExist(colOriSelProduct_, tmpKey) Then
            callerFrm_.RemoveTripControl tmpKey
        End If
    Next
    For X = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(X).TimeTableNo & "_" & colItemAdd_(X).TripNo & "'")
        If isExist(colOriSelProduct_, tmpKey) = False Then
            colOriSelProduct_.Add colItemAdd_(X), tmpKey
            callerFrm_.AddTripControl colItemAdd_(X)
        End If
    Next
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error GoTo ErrD
Dim tmpS As New BWGSearchManager
Dim tmpList As ListItem, tmpTrip As BWGJobDataTimeTable
Dim X%, criStr As String, sqlStr As String, TrailerType As String
    criStr = Replace(Replace(Trim(txtCriteria.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtCriteria.SetFocus
        Exit Sub
    End If
    
    sqlStr = "(WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
    sqlStr = sqlStr & " And (CarRegisID Like '%" & criStr & "%' Or TimeTableNo like '%" & criStr & "%')"
'    sqlStr = SelectWithCompanyID(sqlStr, True)
    Set colSearchResult_ = tmpS.JobDataTimeTableTruckSearch(sqlStr)
    lvResult.ListItems.Clear
    For X = 1 To colSearchResult_.Count
        Set tmpTrip = colSearchResult_(X)
        TrailerType = tmpTrip.TrailerType
        If TrailerType <> "T" Then
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpTrip.TimeTableNo & "_" & tmpTrip.TripNo & "'", tmpTrip.CompanyCode)
            tmpList.SubItems(1) = tmpTrip.TimeTableNo
            tmpList.SubItems(2) = FormatYMD_to_DMY(tmpTrip.WorkDate, "/", "/")
            tmpList.SubItems(3) = tmpTrip.TruckTypeDesc & ", " & tmpTrip.TruckSubTypeDesc
            If Trim(tmpTrip.TrailerTruckNo) = "" Then
                tmpList.SubItems(3) = tmpTrip.TruckTypeDesc & ", " & tmpTrip.TruckSubTypeDesc
                tmpList.SubItems(4) = tmpTrip.CarRegisNo
            Else
                tmpList.SubItems(3) = tmpTrip.TruckTypeDesc & ", " & tmpTrip.TruckSubTypeDesc & " (พ่วง)"
                tmpList.SubItems(4) = tmpTrip.CarRegisNo & " - " & tmpTrip.TrailerTruckNo
            End If
            tmpList.SubItems(5) = tmpTrip.DriverName
            tmpList.SubItems(6) = tmpTrip.FromAumphurName & ", " & tmpTrip.FromProvinceName
            tmpList.SubItems(7) = tmpTrip.ToAumphurName & ", " & tmpTrip.ToProvinceName
            tmpList.SubItems(8) = Format(tmpTrip.VDTranFee, "#,##0.00")
            tmpList.Checked = isExist(lvSelected.ListItems, tmpList.Key)
        End If
        Set tmpList = Nothing
        Set tmpTrip = Nothing
    Next
    If lvResult.ListItems.Count = 0 And Me.Visible Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Label1(2).Caption = "ผลการค้นหา พบ " & lvResult.ListItems.Count & " รายการ"
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
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

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpList As ListItem
Dim X%, selWasteID$
    If item.Checked = True Then
        selWasteID = Trim(item.Key)
        For X = 1 To lvSelected.ListItems.Count
            If Trim(selWasteID) = Trim(lvSelected.ListItems(X).Key) Then
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
            tmpList.SubItems(5) = item.SubItems(5)
            tmpList.SubItems(6) = item.SubItems(6)
            tmpList.SubItems(7) = item.SubItems(7)
            tmpList.SubItems(8) = item.SubItems(8)
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
    Label1(1).Caption = "รายการใบคุมที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Private Sub lvResult_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Pont As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "เลือกทั้งหมด"
        AppendMenu hMenu, MF_STRING, 2, "ไม่เลือกทั้งหมด"

        GetCursorPos Pont
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pont.X, Pont.y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu
        
        Dim i&, item As MSComctlLib.ListItem
        With lvResult
            If ret = 1 Then
                  For i = 1 To .ListItems.Count
                    .ListItems(i).Checked = True
                    Set item = .ListItems(i)
                    lvResult_ItemCheck item
                  Next i
            ElseIf ret = 2 Then
                For i = 1 To .ListItems.Count
                    .ListItems(i).Checked = False
                    Set item = .ListItems(i)
                    lvResult_ItemCheck item
                  Next i
            End If
        End With
    End If
End Sub

Private Sub lvSelected_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim tmpKey As String
    With lvSelected
        If KeyCode = vbKeyDelete And Not .SelectedItem Is Nothing Then
            selResult = MsgBox("ต้องการลบใบคุมนี้ออกจากการทำจ่ายเงินให้ Vendor หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = Trim(.SelectedItem.Key)
                If isExist(colItemDel_, tmpKey) = False Then
                    If isExist(colOriSelProduct_, tmpKey) Then
                        colItemDel_.Add colOriSelProduct_(tmpKey), tmpKey
                    ElseIf isExist(colItemAdd_, tmpKey) Then
                        colItemDel_.Add colItemAdd_(tmpKey), tmpKey
                        colItemAdd_.Remove tmpKey
                    Else
                        MsgBox "ไม่พบใบคุมที่ต้องการลบ", vbExclamation
                    End If
                End If
                If isExist(lvSelected.ListItems, tmpKey) = True Then
                    lvSelected.ListItems.Remove tmpKey
                End If
                If isExist(lvResult.ListItems, tmpKey) Then lvResult.ListItems(tmpKey).Checked = False
            End If
        End If
    End With
End Sub

Public Sub SetCallerForm(frm As Form, selVendor As BWGTransportVendor, curBillHeader As BWGBillPayHeader)
Dim X%, tmpTrip As BWGJobDataTimeTable
Dim tmpList As ListItem
    Set callerFrm_ = frm
    Set curVendor_ = selVendor
    lbCustomer.Caption = selVendor.NameTH
    Set colOriSelProduct_ = frm.GetTripControls
    lvSelected.ListItems.Clear
    Set curBillHeader_ = curBillHeader
    
    If Not colOriSelProduct_ Is Nothing Then
        For X = 1 To colOriSelProduct_.Count
            Set tmpTrip = colOriSelProduct_(X)
            Set tmpList = lvSelected.ListItems.Add(, "'" & tmpTrip.TimeTableNo & "_" & tmpTrip.TripNo & "'", tmpTrip.CompanyCode)
            tmpList.SubItems(1) = tmpTrip.TimeTableNo
            tmpList.SubItems(2) = FormatYMD_to_DMY(tmpTrip.WorkDate, "/", "/")
            If Trim(tmpTrip.TrailerTruckNo) = "" Then
                tmpList.SubItems(3) = tmpTrip.TruckTypeDesc & ", " & tmpTrip.TruckSubTypeDesc
                tmpList.SubItems(4) = tmpTrip.CarRegisNo
            Else
                tmpList.SubItems(3) = tmpTrip.TruckTypeDesc & ", " & tmpTrip.TruckSubTypeDesc & " (พ่วง)"
                tmpList.SubItems(4) = tmpTrip.CarRegisNo & " - " & tmpTrip.TrailerTruckNo
            End If
            tmpList.SubItems(5) = tmpTrip.DriverName
            tmpList.SubItems(6) = tmpTrip.FromAumphurName & ", " & tmpTrip.FromProvinceName
            tmpList.SubItems(7) = tmpTrip.ToAumphurName & ", " & tmpTrip.ToProvinceName
            tmpList.SubItems(8) = Format(tmpTrip.TripTranFee + (tmpTrip.VDEquipFee + tmpTrip.VDWorkerFee + tmpTrip.VDOtherFee + tmpTrip.VDTranFee), "#,##0.00")
            Set tmpList = Nothing
            Set tmpTrip = Nothing
        Next
    End If
    cmdSearch_Click
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub
