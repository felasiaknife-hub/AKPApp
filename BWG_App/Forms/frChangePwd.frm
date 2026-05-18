VERSION 5.00
Begin VB.Form frmChangePwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "แก้ไขรหัสผ่าน"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   3360
      Picture         =   "frChangePwd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   795
      Left            =   2250
      Picture         =   "frChangePwd.frx":057C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1065
   End
   Begin VB.TextBox txtConPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1140
      Width           =   2025
   End
   Begin VB.TextBox txtNewPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2025
   End
   Begin VB.TextBox txtOldPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   300
      Width           =   2025
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยืนยันรหัสผ่านใหม่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   390
      TabIndex        =   7
      Top             =   1200
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผ่านใหม่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   780
      TabIndex        =   6
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผ่านเดิม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   810
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   -360
      Picture         =   "frChangePwd.frx":0665
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   4890
   End
End
Attribute VB_Name = "frmChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curUser_ As BWGCompanyStaff

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim EncObj As RC4Crypt
Dim tmpPwd As String
    If curUser_ Is Nothing Then
        Set curUser_ = CurrentUser
    End If
    If curUser_ Is Nothing Then Exit Sub
    If Trim(txtOldPwd.Text) = "" Then
        MsgBox "กรุณาใส่รหัสผ่านเดิม", vbExclamation
        txtOldPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtNewPwd.Text) = "" Then
        MsgBox "กรุณาใส่รหัสผ่านใหม่", vbExclamation
        txtNewPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtConPwd.Text) = "" Then
        MsgBox "กรุณายืนยันรหัสผ่านใหม่", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtNewPwd.Text) <> Trim(txtConPwd.Text) Then
        MsgBox "การยืนยันรหัสผ่านไม่ถูกต้อง", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
    Set EncObj = New RC4Crypt
    tmpPwd = EncObj.DecryptString(curUser_.LoginPwd, ENC_KEY)
    If tmpPwd <> txtOldPwd.Text Then
        MsgBox "รหัสผ่านเดิมไม่ถูกต้อง กรุณาตรวจสอบ", vbExclamation
        txtOldPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtNewPwd.Text) <> Trim(txtConPwd.Text) Then
        MsgBox "การยืนยันรหัสผ่านไม่ถูกต้อง", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
'    If curUser_ Is Nothing Then
'        CurrentUser.ChangePassword Trim(txtNewPwd.Text)
'    Else
        curUser_.ChangePassword Trim(txtNewPwd.Text)
'    End If
    MsgBox "รหัสผ่านถูกเปลี่ยนเรียบร้อยแล้ว กรุณาใช้รหัสผ่านนี้ในการเข้าระบบครั้งต่อไป", vbInformation
    txtOldPwd.Text = ""
    txtNewPwd.Text = ""
    txtConPwd.Text = ""
    Unload Me
End Sub

Public Sub SetShowUser(selUser As BWGCompanyStaff)
    Set curUser_ = selUser
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set curUser_ = Nothing
End Sub
