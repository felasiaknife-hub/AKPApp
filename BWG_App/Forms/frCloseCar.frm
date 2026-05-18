VERSION 5.00
Begin VB.Form frmCloseCar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ปิดรถ - F009"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   795
      Left            =   5640
      Picture         =   "frCloseCar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2130
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6750
      Picture         =   "frCloseCar.frx":00E9
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2130
      Width           =   1065
   End
   Begin VB.TextBox txtConPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1650
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1410
      Width           =   5565
   End
   Begin WasteManagment.ctlDate ctlDate1 
      Height          =   315
      Left            =   1650
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate ctlDate2 
      Height          =   315
      Left            =   5430
      TabIndex        =   4
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate ctlDate3 
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Top             =   990
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlDate ctlDate4 
      Height          =   315
      Left            =   5430
      TabIndex        =   8
      Top             =   990
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8070
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "หมายเหตุ"
      Height          =   195
      Index           =   4
      Left            =   900
      TabIndex        =   10
      Top             =   1470
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "วันที่ปิดรถ"
      Height          =   195
      Index           =   3
      Left            =   4620
      TabIndex        =   7
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "วันที่รถเข้าศุนย์"
      Height          =   195
      Index           =   2
      Left            =   450
      TabIndex        =   5
      Top             =   1050
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "วันที่รถออกจากโรงงาน"
      Height          =   195
      Index           =   1
      Left            =   3750
      TabIndex        =   3
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "วันที่รถเข้าโรงงาน"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Width           =   1260
   End
   Begin VB.Label lbCarDesc 
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   7305
   End
End
Attribute VB_Name = "frmCloseCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If MsgBox("ท่านต้องการปิดรถคันนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
    
        MsgBox "ทำการปิดรถเรียบร้อยแล้ว", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
End Sub
