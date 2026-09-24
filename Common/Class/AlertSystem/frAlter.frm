VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ตั้งเวลาเตือน"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "เคลียร์หน้าจอ"
      Height          =   375
      Left            =   3690
      TabIndex        =   8
      Top             =   5820
      Width           =   1395
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   375
      Left            =   6450
      TabIndex        =   10
      Top             =   5820
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   5820
      Width           =   1455
   End
   Begin VB.TextBox txtMsg 
      Height          =   795
      Left            =   1950
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "4"
      Top             =   1320
      Width           =   5865
   End
   Begin VB.CheckBox chkWornClose 
      Caption         =   "ทำการเตือนก่อนปิดโปรแกรม"
      Height          =   285
      Left            =   4313
      TabIndex        =   4
      Top             =   720
      Width           =   2385
   End
   Begin VB.CheckBox chkWornOpen 
      Caption         =   "ทำการเตือนเมื่อเริ่มใช้โปรแกรม"
      Height          =   285
      Left            =   1433
      TabIndex        =   3
      Top             =   720
      Width           =   2625
   End
   Begin VB.CheckBox chkEnd 
      Caption         =   "จนกว่าผู้ใช้จะสั่งหยุด"
      Height          =   345
      Left            =   6090
      TabIndex        =   2
      Top             =   120
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   56754179
      CurrentDate     =   37327.0229166667
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   3990
      TabIndex        =   1
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   56754179
      CurrentDate     =   37327.0229166667
   End
   Begin MSComctlLib.ListView lvAlertList 
      Height          =   3375
      Left            =   30
      TabIndex        =   6
      Top             =   2310
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "วันที่เตือน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ข้อความเตือน"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "สถานะ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbCount 
      Height          =   225
      Left            =   60
      TabIndex        =   15
      Top             =   5880
      Width           =   1995
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   2
      Left            =   -60
      Top             =   2190
      Width           =   9825
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      Caption         =   "ข้อความเตือนที่จะแสดง"
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   14
      Tag             =   "PROD_DESC"
      Top             =   1380
      Width           =   1605
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -90
      Top             =   1170
      Width           =   9825
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   0
      Left            =   -90
      Top             =   540
      Width           =   9825
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      Caption         =   "หรือ"
      Height          =   195
      Index           =   2
      Left            =   5730
      TabIndex        =   13
      Tag             =   "PROD_DESC"
      Top             =   180
      Width           =   285
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   0
      Left            =   3420
      TabIndex        =   12
      Tag             =   "PROD_DESC"
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      Caption         =   "เริ่มเตือนตั้งแต่วันที่"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   11
      Tag             =   "PROD_DESC"
      Top             =   180
      Width           =   1380
   End
End
Attribute VB_Name = "frAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curAlerts_ As Collection
Dim curAlert_ As AWSAlert
Dim isChanged_ As Boolean

Private Sub chkEnd_Validate(Cancel As Boolean)
    If curAlert_ Is Nothing Then Exit Sub
    curAlert_.ManualStop = chkEnd.Value
End Sub

Private Sub chkWornClose_Validate(Cancel As Boolean)
    If curAlert_ Is Nothing Then Exit Sub
    curAlert_.AlertAtEnd = chkWornClose.Value
End Sub

Private Sub chkWornOpen_Validate(Cancel As Boolean)
    If curAlert_ Is Nothing Then Exit Sub
    curAlert_.AlertAtStart = chkWornOpen.Value
End Sub

Private Sub cmdClear_Click()
Dim tmpVal As Integer
    If curAlert_ Is Nothing Then Exit Sub

    If curAlert_.isChanged Then
        tmpVal = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บหรือไม่", vbQuestion + vbYesNoCancel)
        If tmpVal = vbYes Then
            cmdSave_Click
            If curAlert_.isChanged Then Exit Sub
        ElseIf tmpVal = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewAlert
    txtMsg.SetFocus
End Sub

Private Sub cmdClose_Click()
    If isChanged_ = True Then
        frMain.LoadAlertList
    End If
    Unload Me
End Sub
'
Private Sub cmdDel_Click()
    If curAlert_ Is Nothing Then
        MsgBox "กรุณาเลือกรายการเตือนที่ต้องการลบ", vbExclamation
        Exit Sub
    End If

    If MsgBox("ต้องการลบข้อมูลการเตือน '" & curAlert_.MessageSingleLine & "' ออกจากฐานข้อมูลหรือไม่", vbYesNo + vbQuestion) = vbYes Then
        curAlert_.Delete
        ClearScreen
        If isExist(lvAlertList.ListItems, "'" & curAlert_.AlertID & "'") Then
            lvAlertList.ListItems.Remove "'" & curAlert_.AlertID & "'"
            curAlerts_.Remove "'" & curAlert_.AlertID & "'"
        End If
        CreateNewAlert
        isChanged_ = True
        txtMsg.SetFocus
    End If
    lbCount.Caption = lvAlertList.ListItems.Count & " รายการ"
End Sub

Private Sub cmdSave_Click()
Dim tmpItem As listItem
    If curAlert_ Is Nothing Then Exit Sub
    If Trim(txtMsg.Text) = "" Then
        MsgBox "กรุณาใส่ข้อความเตือนที่จะแสดง", vbExclamation
        txtMsg.SetFocus
        Exit Sub
    End If

    dtStart_Validate False
    dtEnd_Validate False
    curAlert_.Save
    MsgBox "เก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    If isExist(lvAlertList.ListItems, "'" & curAlert_.AlertID & "'") Then
        Set tmpItem = lvAlertList.ListItems("'" & curAlert_.AlertID & "'")
        tmpItem.Text = FormatYMD_to_DMY(curAlert_.startDate, "/", "/")
    Else
        Set tmpItem = lvAlertList.ListItems.Add(, "'" & curAlert_.AlertID & "'", FormatYMD_to_DMY(curAlert_.startDate, "/", "/"))
        If curAlerts_ Is Nothing Then Set curAlerts_ = New Collection
        curAlerts_.Add curAlert_, "'" & curAlert_.AlertID & "'"
    End If
    tmpItem.SubItems(1) = curAlert_.MessageSingleLine
    tmpItem.SubItems(2) = curAlert_.ActiveStatus
    If curAlert_.isActive = 0 Then
        tmpItem.ForeColor = &H808080
        tmpItem.ListSubItems(1).ForeColor = &H808080
        tmpItem.ListSubItems(2).ForeColor = &H808080
    End If

    Set tmpItem = Nothing
    ClearScreen
    CreateNewAlert
    txtMsg.SetFocus
    isChanged_ = True
    lbCount.Caption = lvAlertList.ListItems.Count & " รายการ"
End Sub

Private Sub dtEnd_Validate(Cancel As Boolean)
    If curAlert_ Is Nothing Then Exit Sub
    curAlert_.endDate = FormatYMD_Thai(dtEnd.Day, dtEnd.Month, dtEnd.Year, "/")
End Sub

Private Sub dtStart_Validate(Cancel As Boolean)
    If curAlert_ Is Nothing Then Exit Sub
    curAlert_.startDate = FormatYMD_Thai(dtStart.Day, dtStart.Month, dtStart.Year, "/")
End Sub

Private Sub Form_Load()
    ShowAlertList
    dtStart.Value = Date
    dtEnd.Value = DateSerial(Year(Date), Month(Date) + 1, 1)
    CreateNewAlert
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim tmpVal As Integer
    If curAlert_ Is Nothing Then Exit Sub
    If curAlert_.isChanged Then
        tmpVal = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If tmpVal = vbYes Then
            cmdSave_Click
            If curAlert_.isChanged Then Cancel = 1
        ElseIf tmpVal = vbCancel Then
            Cancel = 1
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set curAlerts_ = Nothing
    Set curAlert_ = Nothing
End Sub

Private Sub CreateNewAlert()
    Set curAlert_ = Nothing
    Set curAlert_ = New AWSAlert
    With curAlert_
        .AlertID = getNewProductID(modPrefix.AlertManualPrefix)
        .startDate = FormatYMD_Thai(dtStart.Day, dtStart.Month, dtStart.Year, "/")
        .endDate = FormatYMD_Thai(dtEnd.Day, dtEnd.Month, dtEnd.Year, "/")
        .isChanged = False
    End With
End Sub

Private Sub ClearScreen()
    dtStart.Value = Date
    dtEnd.Value = Date
    chkEnd.Value = 0
    Me.chkWornClose.Value = 0
    Me.chkWornOpen.Value = 0
    txtMsg.Text = ""
End Sub

Private Sub ShowAlertData()

    If curAlert_ Is Nothing Then Exit Sub

    With curAlert_
        If IsDate(.startDate) Then dtStart.Value = .startDate
        If IsDate(.endDate) Then dtEnd.Value = .endDate
        chkEnd.Value = .ManualStop
        chkWornClose.Value = .AlertAtEnd
        chkWornOpen.Value = .AlertAtStart
        txtMsg.Text = .Message
        .isChanged = False
    End With
End Sub

Private Sub lvAlertList_ItemClick(ByVal Item As MSComctlLib.listItem)
    If curAlerts_ Is Nothing Then Exit Sub
    If lvAlertList.ListItems Is Nothing Then Exit Sub

    Set curAlert_ = Nothing
    Set curAlert_ = curAlerts_(Item.Key)
    ShowAlertData
End Sub

Private Sub lvAlertList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cmdDel_Click
    End If
End Sub

Private Sub txtMsg_Validate(Cancel As Boolean)
    If curAlert_ Is Nothing Then Exit Sub
    curAlert_.Message = txtMsg.Text
End Sub

Private Sub ShowAlertList()
Dim x%
Dim tmpAlert As AWSAlert
Dim listItem As listItem
    Set curAlerts_ = AlertManager.getAllAlerts  'AlertManager.getActiveAlerts
    If curAlerts_ Is Nothing Then Exit Sub
    For x = 1 To curAlerts_.Count
        Set tmpAlert = curAlerts_(x)
        Set listItem = lvAlertList.ListItems.Add(, "'" & tmpAlert.AlertID & "'", FormatYMD_to_DMY(tmpAlert.startDate, "/", "/"))
        listItem.SubItems(1) = TrimString(tmpAlert.MessageSingleLine, 50)
        listItem.SubItems(2) = tmpAlert.ActiveStatus
        If tmpAlert.isActive = 0 Then
            listItem.ForeColor = &H808080
            listItem.ListSubItems(1).ForeColor = &H808080
            listItem.ListSubItems(2).ForeColor = &H808080
        End If

        Set listItem = Nothing
        Set tmpAlert = Nothing
    Next
    lbCount.Caption = lvAlertList.ListItems.Count & " รายการ"
End Sub
