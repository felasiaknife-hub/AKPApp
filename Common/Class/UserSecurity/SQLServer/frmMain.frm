VERSION 5.00
Begin VB.MDIForm frMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7860
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   6390
      Left            =   0
      ScaleHeight     =   6330
      ScaleWidth      =   3030
      TabIndex        =   0
      Top             =   0
      Width           =   3090
      Begin VB.CommandButton Command1 
         Caption         =   "User Manager for Administrator"
         Height          =   1035
         Left            =   0
         Picture         =   "frmMain.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1170
         Width           =   2985
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change Password for Individual"
         Height          =   1035
         Left            =   0
         Picture         =   "frmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3450
         Width           =   2985
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Login"
         Height          =   1035
         Left            =   0
         Picture         =   "frmMain.frx":09A1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   2985
      End
      Begin VB.CommandButton Command4 
         Caption         =   "User Group"
         Height          =   1035
         Left            =   0
         Picture         =   "frmMain.frx":0A78
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4590
         Width           =   2985
      End
      Begin VB.CommandButton Command5 
         Caption         =   "AAdd New User"
         Height          =   1035
         Left            =   0
         Picture         =   "frmMain.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2310
         Width           =   2985
      End
   End
   Begin VB.Menu mnu_PopUser 
      Caption         =   "aaaa"
      Begin VB.Menu sdfsdf 
         Caption         =   "sfsd"
      End
      Begin VB.Menu erer 
         Caption         =   "erere"
      End
   End
End
Attribute VB_Name = "frMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    'frUserProfile.Show 1
    If Not CurrentUser Is Nothing Then
        If CurrentUser.VerifyAccessRole("frUser", AccessRole) = False Then
            MsgBox "ไม่มีสิทธิ์ใช้ระบบนี้"
            Exit Sub
        End If
    End If
    frUserList.Show
End Sub

Private Sub Command2_Click()
    frChangePwd.Show 1
End Sub

Private Sub Command3_Click()
    frLogin.Show 1
End Sub

Private Sub Command4_Click()
    frUserGroup.Show 1
End Sub

Private Sub Command5_Click()
    If Not CurrentUser Is Nothing Then
        If CurrentUser.VerifyAccessRole("frUser", AccessRole) = False Then
            MsgBox "ไม่มีสิทธิ์ใช้ระบบนี้"
            Exit Sub
        End If
    End If
    
    frUserProfile.Show 1
End Sub

