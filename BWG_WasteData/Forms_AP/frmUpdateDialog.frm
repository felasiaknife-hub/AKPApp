VERSION 5.00
Begin VB.Form frmUpdateDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Dialog"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "อัพเดทโปรแกรม"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   -120
      TabIndex        =   2
      Top             =   -120
      Width           =   6255
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "มีโปรแกรมใหม่ สำหรับอัพเดท กรุณาคลิกที่ปุ่ม อัพเดทโปรแกรม โปรแกรมจะถูกอัพเดทอัตโนมัติ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   75
      Left            =   -120
      ScaleHeight     =   15
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   2520
      Width           =   6135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "กรุณาปิดโปรแกรม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1140
      TabIndex        =   4
      Top             =   1620
      Width           =   3675
   End
End
Attribute VB_Name = "frmUpdateDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdUpdate_Click()
Dim Frm As Form
    For Each Frm In Forms
        Unload Frm
    Next
    ShellExecute 0, vbNullString, App.Path & "\BWGUpdate.bat", vbNullString, vbNullString, vbNormalFocus
End Sub


