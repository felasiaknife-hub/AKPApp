VERSION 5.00
Begin VB.Form frmJobDetFail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ข้อผิดพลากในการขนกากของเสีย - F040"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "ไม่มีของ"
      Height          =   255
      Left            =   1380
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2940
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2940
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4650
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2940
      Width           =   945
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   6135
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1290
      Width           =   6105
   End
   Begin VB.TextBox txtWasteName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   900
      Width           =   6105
   End
   Begin VB.TextBox txtCustName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   510
      Width           =   6105
   End
   Begin VB.TextBox txtManifestNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   120
      Width           =   1755
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7800
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำอธิบาย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   540
      TabIndex        =   12
      Top             =   2040
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "สาเหตุเกิดจาก"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   11
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กากของเสีย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   390
      TabIndex        =   10
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   810
      TabIndex        =   9
      Top             =   570
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งงาน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   180
      Width           =   990
   End
End
Attribute VB_Name = "frmJobDetFail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim selJobDet_ As BWGJobDataDetail




Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If MsgBox("ท่านต้องการลบข้อมูลข้อผิดพลาดในการขนส่งของกากของเสียนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        'DBConnExc "Delete tbJobFail WHERE JobDetID='" & selJobDet_.ID & "'"
        cboType.ListIndex = -1
        txtRem.Text = ""
        cboType.SetFocus
        selJobDet_.FailRemark = ""
        selJobDet_.FailType = ""
        selJobDet_.FailCheck = 0
    End If
End Sub

Private Sub cmdSave_Click()
    If cboType.ListIndex = -1 Then
        MsgBox "กรุณาระบุสาเหตุที่ทำให้เกิดข้อผิดพลาดในการขนกากของเสีย", vbExclamation
        cboType.SetFocus
        Exit Sub
    End If
    If isChanged Then
        With selJobDet_
            Select Case cboType.ListIndex
                Case 0
                    .FailType = "A"
                Case 1
                    .FailType = "B"
                Case 2
                    .FailType = "C"
            End Select
            .FailRemark = Trim(txtRem.Text)
            .FailCheck = Check1.Value
        End With
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With cboType
        .AddItem FailTypeDesc("A")
        .AddItem FailTypeDesc("B")
        .AddItem FailTypeDesc("C")
    End With
End Sub

Public Sub SetJobDetail(selDet As BWGJobDataDetail, docNo As String, custName As String)
    Set selJobDet_ = selDet
    With selJobDet_
        txtManifestNo.Text = docNo
        txtCustName.Text = custName
        txtWasteName.Text = .WasteName
        SetTextToCombo modMain.FailTypeDesc(.FailType), cboType
        txtRem.Text = .FailRemark
        Check1.Value = .FailCheck
    End With
End Sub

Private Function isChanged() As Boolean
Dim tmpType As String
    Select Case cboType.ListIndex
        Case 0
            tmpType = "A"
        Case 1
            tmpType = "B"
        Case 2
            tmpType = "C"
    End Select
    With selJobDet_
        If Trim(.FailRemark) <> Trim(txtRem.Text) Or Trim(.FailType) <> Trim(tmpType) Or .FailCheck <> Check1.Value Then
            isChanged = True
        Else
            isChanged = False
        End If
    End With
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub
