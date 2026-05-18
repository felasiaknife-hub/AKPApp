VERSION 5.00
Begin VB.Form frmLicenseStartDate 
   Caption         =   "กำหนดวันเริ่มต้นสัญญา"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin WasteDataForm.ctlDate LStartDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin WasteDataForm.ctlDate LEndDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "วันที่หมดอายุ :"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "วันที่เริ่มสัญญา :"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmLicenseStartDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmWasteData.LicenseStartDate = LStartDate.ValueYMD
    frmWasteData.LicenseExpDate = LEndDate.ValueYMD
    frmWasteData.setLicenseStartDate
    Unload Me
End Sub
Public Sub SetCallerForm(selFrm As Form)
    Set callerForm_ = selFrm
End Sub
Private Sub Form_Load()
      LStartDate.ValueYMD = frmWasteData.LicenseStartDate
      LEndDate.ValueYMD = frmWasteData.LicenseExpDate
End Sub
