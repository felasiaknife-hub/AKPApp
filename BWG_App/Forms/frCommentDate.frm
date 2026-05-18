VERSION 5.00
Begin VB.Form frmCommentDate 
   Caption         =   "ระบุวันที่"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin WasteManagment.ctlDate ctlDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ตกลง"
      Height          =   435
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   435
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "ระบุวันที่ลูกค้าหลุด :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmCommentDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmCustomer.lbDate.Caption = ctlDate.ValueYMD
    frmCustomer.AddCommentDate
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
