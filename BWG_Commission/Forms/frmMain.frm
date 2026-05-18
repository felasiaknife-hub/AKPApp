VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "รายงาน"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   -60
      TabIndex        =   6
      Top             =   2280
      Width           =   5715
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   -60
      TabIndex        =   4
      Top             =   -120
      Width           =   6015
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "รายงานเกี่ยวกับค่าคอมฯ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdReComEx 
      Caption         =   "การตัดจ่ายค่าคอมนอก"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdReComIn 
      Caption         =   "การตัดจ่ายค่าคอมใน"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdComEx 
      Caption         =   "การตั้งค่าคอมนอก"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   900
      Width           =   2175
   End
   Begin VB.CommandButton cmdComIn 
      Caption         =   "การตั้งค่าคอมใน"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComEx_Click()
    frmComExt1.Show
End Sub

Private Sub cmdReComEx_Click()
    frmComExt2.Show
End Sub

Private Sub cmdComIn_Click()
    frmComInt1.Show
End Sub

Private Sub cmdReComIn_Click()
    frmComInt2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

