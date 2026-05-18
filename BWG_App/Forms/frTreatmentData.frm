VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTreatmentData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F130"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRateWeight 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      MaxLength       =   80
      TabIndex        =   2
      Top             =   870
      Width           =   1035
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4860
      MaxLength       =   80
      TabIndex        =   1
      Top             =   870
      Width           =   1035
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5190
      Width           =   1035
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      MaxLength       =   80
      TabIndex        =   0
      Top             =   870
      Width           =   2505
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5190
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5190
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3825
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6747
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "วิธีการกำจัด"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "อัตราค่ากำจัด/เที่ยว"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "อัตราค่ากำจัด/ตัน"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "บาท/ตัน"
      Height          =   195
      Index           =   4
      Left            =   7950
      TabIndex        =   11
      Top             =   930
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "บาท/เที่ยว"
      Height          =   195
      Index           =   2
      Left            =   5970
      TabIndex        =   10
      Top             =   930
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "อัตราค่ากำจัด"
      Height          =   195
      Index           =   0
      Left            =   3870
      TabIndex        =   9
      Top             =   930
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5640
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   5790
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูลวิธีการกำจัดของเสีย"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   225
      TabIndex        =   8
      Top             =   0
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "วิธีการกำจัด"
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   7
      Top             =   930
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frTreatmentData.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8730
   End
End
Attribute VB_Name = "frmTreatmentData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colTreatment_ As Collection
Dim curTreatment_ As BWGTreatmentData
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Sub cmdClose_Click()
Dim aa
    If isChanged Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนปิดหรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If isChanged Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdNew_Click()
Dim aa
    If isChanged Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If isChanged Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewTreatment
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub cmdSave_Click()
Dim sqlStr As String
Dim tmpList As ListItem
Dim tmpCon As ADODB.Connection
Dim tmpStr As String
Dim treatRate As Double, treatRateW As Double
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If curTreatment_ Is Nothing Then Exit Sub
    If Trim(txtShortName.Text) = vbNullString Then
        MsgBox "กรุณาใส่ชื่อหน่วยก่อนเก็บข้อมูล", vbExclamation
        txtShortName.SetFocus
        Exit Sub
    End If
    If Trim(txtRate.Text) <> "" And IsNumeric(txtRate.Text) = False Then
        MsgBox "กรุณาใส่อัตราค่ากำจัดเป็นตัวเลขเท่านั้น", vbExclamation
        txtRate.SetFocus
        Exit Sub
    End If
    If Trim(txtRateWeight.Text) <> "" And IsNumeric(txtRateWeight.Text) = False Then
        MsgBox "กรุณาใส่อัตราค่ากำจัดเป็นตัวเลขเท่านั้น", vbExclamation
        txtRateWeight.SetFocus
        Exit Sub
    End If

    If IsNumeric(txtRate.Text) Then
        treatRate = txtRate.Text
    Else
        treatRate = 0
    End If
    If IsNumeric(txtRateWeight.Text) Then
        treatRateW = txtRateWeight.Text
    Else
        treatRateW = 0
    End If
    
    If isChanged Then
        If Trim(curTreatment_.ID) = "" Then
            curTreatment_.ID = GetGUID
            DBConnExc "Insert into tbTreatmentData (TreatmentID,TreatmentDesc,TreatmentRate,TreatmentUnit,TreatmentRateWeight,TreatmentUnitWeight) Values ('" & Trim(curTreatment_.ID) & "','" & (txtShortName.Text) & "'," & treatRate & ",'บาท/เที่ยว'," & treatRateW & ",'บาท/ตัน')"
            Set tmpList = lvList.ListItems.Add(, "'" & curTreatment_.ID & "'", txtShortName.Text)
            colTreatment_.Add curTreatment_, "'" & curTreatment_.ID & "'"
        Else
            DBConnExc "Update tbTreatmentData Set TreatmentDesc='" & Trim(txtShortName.Text) & "',TreatmentRate=" & treatRate & ",TreatmentUnit='บาท/เที่ยว',TreatmentRateWeight=" & treatRateW & ",TreatmentUnitWeight='บาท/ตัน' WHERE TreatmentID='" & curTreatment_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curTreatment_.ID & "'")
        End If
        curTreatment_.treatmentDesc = Trim(txtShortName.Text)
        curTreatment_.TreatmentRate = txtRate.Text
        curTreatment_.TreatmentUnit = "บาท/เที่ยว"
        curTreatment_.TreatmentRateWeight = txtRateWeight.Text
        curTreatment_.TreatmentUnitWeight = "บาท/ตัน"
        tmpList.Text = Trim(txtShortName.Text)
        tmpList.SubItems(1) = Format(curTreatment_.TreatmentRate, "#,##0.00") & " " & curTreatment_.TreatmentUnit
        tmpList.SubItems(2) = Format(curTreatment_.TreatmentRateWeight, "#,##0.00") & " " & curTreatment_.TreatmentUnitWeight
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewTreatment
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim searchM As New BWGSearchManager
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdNew.Picture = .picNew.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Set colTreatment_ = searchM.TreatmentDataSearch
    ShowTreatmentList
    CreateNewTreatment
End Sub

Private Sub ShowTreatmentList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGTreatmentData
    If colTreatment_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colTreatment_.Count
        Set tmpItem = colTreatment_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.treatmentDesc)
        tmpList.SubItems(1) = Format(tmpItem.TreatmentRate, "#,##0.00") & " " & tmpItem.TreatmentUnit
        tmpList.SubItems(2) = Format(tmpItem.TreatmentRateWeight, "#,##0.00") & " " & tmpItem.TreatmentUnitWeight
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.SetDataTreatment colTreatment_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curTreatment_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colTreatment_, tmpKey) Then
        Set curTreatment_ = colTreatment_(tmpKey)
        txtShortName.Text = curTreatment_.treatmentDesc
        txtRate.Text = curTreatment_.TreatmentRate
        txtRateWeight.Text = curTreatment_.TreatmentRateWeight
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewTreatment()
    Set curTreatment_ = New BWGTreatmentData
    curTreatment_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
    txtRate.Text = ""
    txtRateWeight.Text = ""
End Sub

Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpKey As String
Dim tmpStr As String
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, DeleteRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If lvList.SelectedItem Is Nothing Then Exit Sub
    If curTreatment_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curTreatment_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curTreatment_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colTreatment_, tmpKey) Then colTreatment_.Remove tmpKey
                CreateNewTreatment
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Function isChanged() As Boolean
Dim tmpVal As Double, tmpValW As Double
    If IsNumeric(txtRate.Text) Then
        tmpVal = txtRate.Text
    Else
        tmpVal = 0
    End If
    If IsNumeric(txtRateWeight.Text) Then
        tmpValW = txtRateWeight.Text
    Else
        tmpValW = 0
    End If
    
    If Trim(txtShortName.Text) <> Trim(curTreatment_.treatmentDesc) Or tmpVal <> curTreatment_.TreatmentRate Or tmpValW <> curTreatment_.TreatmentRateWeight Then
        isChanged = True
    Else
        isChanged = False
    End If
End Function
