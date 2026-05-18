VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmPromist 
   BackColor       =   &H8000000E&
   Caption         =   "บันทึกข้อตกลงใบเสนอราคา"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancles 
      Caption         =   "ยกเลิก"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtPromies 
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   1980
      Width           =   10275
   End
   Begin DateCtl.Date StartDate 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin DateCtl.Date EndDate 
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label4 
      Caption         =   "อ้างอิงจากใบเสนอราคาเลขที่"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   2235
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "บันทึกข้อตกลงใบเสนอราคา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   480
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   0
      Picture         =   "frmPromist.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label2 
      Caption         =   "สิ้นสุดวันที่ :"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "เริ่มวันที่ :"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmPromist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myQuotation_ As BWGQuotationHeader
Private Sub CmdCancles_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
   frmQuotation.QuotationPromies = txtPromies.Text
   frmQuotation.QuotationStartDatePromies = StartDate.ValueYMD
   frmQuotation.QuotationEndDatePromies = EndDate.ValueYMD
    Unload Me
End Sub

Private Sub Form_Load()
   txtPromies.Text = frmQuotation.QuotationPromies
   If frmQuotation.QuotationStartDatePromies <> "" Then
        StartDate.ValueYMD = frmQuotation.QuotationStartDatePromies
   End If
   If frmQuotation.QuotationEndDatePromies <> "" Then
        EndDate.ValueYMD = frmQuotation.QuotationEndDatePromies
   End If
End Sub
Private Sub StartDate_Validate(Cancel As Boolean)
   frmQuotation.frmPromiesChang = True
End Sub
Private Sub EndDate_Validate(Cancel As Boolean)
    frmQuotation.frmPromiesChang = True
End Sub

Private Sub txtPromies_Validate(Cancel As Boolean)
    frmQuotation.frmPromiesChang = True
End Sub
