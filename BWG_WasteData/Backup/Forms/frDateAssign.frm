VERSION 5.00
Begin VB.Form frmDateAssign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WasteDataForm.ctlDate dtSelDate 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   345
      Left            =   2648
      TabIndex        =   1
      Top             =   1050
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Height          =   345
      Left            =   1208
      TabIndex        =   0
      Top             =   1050
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "frmDateAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form

Private Sub Command1_Click()
    callerForm_.SetDate dtSelDate.ValueYMD
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Public Sub SetCallerForm(selFrm As Form)
    Set callerForm_ = selFrm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

