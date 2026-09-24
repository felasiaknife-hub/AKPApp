VERSION 5.00
Begin VB.Form frLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1710
      TabIndex        =   0
      Top             =   390
      Width           =   2025
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1710
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   810
      Width           =   2025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "เข้าระบบ"
      Default         =   -1  'True
      Height          =   675
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1530
      Width           =   1125
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "ยกเลิก"
      Height          =   675
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1530
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อที่ใช้เข้าระบบ"
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
      Left            =   510
      TabIndex        =   5
      Top             =   450
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผ่าน"
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
      Left            =   990
      TabIndex        =   4
      Top             =   870
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   -30
      X2              =   4650
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   -420
      Picture         =   "frLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4890
   End
End
Attribute VB_Name = "frLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim errCount_ As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim tmpUser As AWSUser
Dim isValid As Boolean
    If Trim(txtUserName.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อที่ใช้ในการเข้าระบบ", vbExclamation
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "กรุณาใส่รหัสผ่าน", vbExclamation
        txtPassword.SetFocus
        Exit Sub
    End If
    Set CurrentUser = modUserSecurity.ProcessLogin(Trim(txtUserName.Text), Trim(txtPassword.Text))
    If Not CurrentUser Is Nothing Then
        Unload Me
        'MsgBox "Login Succeed"
        frMain.Show
    Else
        errCount_ = errCount_ + 1
        If errCount_ < 3 Then
            MsgBox "ข้อมูลการเข้าระบบไม่ถูกต้อง กรุณาตรวจสอบชื่อในการเข้าระบบและรหัสผ่าน", vbExclamation
            txtUserName.SetFocus
            Exit Sub
        Else
            MsgBox "ข้อมูลการเข้าระบบไม่ถูกต้องติดต่อกันครบ 3 ครั้ง โปรแกรมไม่สามารถทำงานต่อได้", vbExclamation
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    errCount_ = 0
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
