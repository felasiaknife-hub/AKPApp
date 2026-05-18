VERSION 5.00
Begin VB.Form frmComRateSup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "อัตราค่าคอมมิชชั่นของ Supervisor - F020"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5730
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5730
      Width           =   1035
   End
   Begin WasteManagment.ctlGrid grComRate 
      Height          =   5625
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   9922
   End
   Begin VB.Shape Shape1 
      Height          =   5685
      Left            =   30
      Top             =   0
      Width           =   6165
   End
End
Attribute VB_Name = "frmComRateSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colComRate_ As Collection
Dim isChanged_ As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim x%
    If isChanged_ Then
        DBConnExc "Delete ComRateSup"
        For x = 1 To colComRate_.Count
            colComRate_(x).Save
        Next
    End If
    isChanged_ = False
    MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
End Sub

Private Sub Form_Load()
    isChanged_ = False
    Set colComRate_ = New Collection
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    
    With grComRate.getGridObj
        .Cols = 4
        .ColWidth(0) = 0
        .ColWidth(1) = 2040
        .ColWidth(2) = 2040
        .ColWidth(3) = 1755
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .Rows = 50
        grComRate.GridMoveType 1
        .TextMatrix(0, 1) = "อัตรายอดขายเริ่มต้น (%)"
        .TextMatrix(0, 2) = "อัตรายอดขายสิ้นสุด (%)"
        .TextMatrix(0, 3) = "อัตราค่าคอมฯ (%)"
    End With
    ShowComRate
End Sub

Private Sub ShowComRate()
Dim tmpS As New BWGSearchManager
Dim x%, tmpCom As BWGComRateSup
    Set colComRate_ = tmpS.ComRateSupSearch
    grComRate.ClearAllData
    With grComRate.getGridObj
        .Rows = 50
        For x = 1 To colComRate_.Count
            Set tmpCom = colComRate_(x)
            .TextMatrix(x, 0) = tmpCom.ID
            .TextMatrix(x, 1) = Format(tmpCom.StartPC, "0.00")
            .TextMatrix(x, 2) = Format(tmpCom.EndPC, "0.00")
            .TextMatrix(x, 3) = Format(tmpCom.ComRate, "0.00")
            Set tmpCom = Nothing
        Next
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If isChanged_ Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If isChanged_ Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    isChanged_ = False
    Set colComRate_ = Nothing
End Sub

Private Sub grComRate_BeforeUpdateValue(newValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If IsNumeric(newValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grComRate_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpKey As String
    With grComRate.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            If MsgBox("ท่านต้องการลบข้อมูลอัตราค่าคอมมิชชั่นนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                If isExist(colComRate_, tmpKey) Then colComRate_.Remove tmpKey
                .RemoveItem .row
                isChanged_ = True
            End If
        End If
    End With
End Sub

Private Sub grComRate_RolColChanged(row As Integer, col As Integer)
    With grComRate.getGridObj
        grComRate.ColEnabled(col) = Trim(.TextMatrix(row - 1, 1)) <> "" And Trim(.TextMatrix(row - 1, 2)) <> "" And Trim(.TextMatrix(row - 1, 3)) <> ""
    End With
End Sub

Private Sub grComRate_UpdateValue(row As Integer, col As Integer)
Dim tmpID As String
Dim tmpVal As String
Dim tmpRate As BWGComRateSup
    With grComRate.getGridObj
        If Trim(.TextMatrix(row, 0)) <> "" Then
            tmpID = Trim(.TextMatrix(row, 0))
        Else
            tmpID = GetGUID
        End If
        tmpVal = Trim(.TextMatrix(row, col))
        If isExist(colComRate_, "'" & tmpID & "'") Then
            Set tmpRate = colComRate_("'" & tmpID & "'")
        Else
            Set tmpRate = New BWGComRateSup
            tmpRate.ID = tmpID
            .TextMatrix(row, 0) = tmpID
            colComRate_.Add tmpRate, "'" & tmpID & "'"
        End If
    End With
    Select Case col
        Case 1
            If IsNumeric(Trim(tmpVal)) Then
                tmpRate.StartPC = Trim(tmpVal)
            Else
                tmpRate.StartPC = 0
            End If
        Case 2
            If IsNumeric(Trim(tmpVal)) Then
                tmpRate.EndPC = Trim(tmpVal)
            Else
                tmpRate.EndPC = 0
            End If
        Case 3
            If IsNumeric(Trim(tmpVal)) Then
                tmpRate.ComRate = Trim(tmpVal)
            Else
                tmpRate.ComRate = 0
            End If
    End Select
    Set tmpRate = Nothing
    isChanged_ = True
End Sub
