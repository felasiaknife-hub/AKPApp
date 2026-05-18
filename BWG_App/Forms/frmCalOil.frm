VERSION 5.00
Begin VB.Form frmCalOil 
   Caption         =   "frmCalOil"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtA 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   3825
   End
   Begin VB.TextBox TxtX 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   3825
   End
   Begin VB.TextBox TxtB 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   3825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   7215
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   795
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton cmdCal 
         Caption         =   "คำนวน"
         Height          =   795
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   23.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   1440
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "ผลรับ"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "อัตราค่าน้ำมันต่อลิตร"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "ปริมาณน้ำมันที่ใช้"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อัตราค่าบริการขนส่งของรถ"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   720
      X2              =   6420
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   480
      X2              =   6120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานระยะทางการขนกากของรถแต่ละคัน (TRN10A-1)"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4890
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmCalOil.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmCalOil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCal_Click()
        Dim a As Double
        Dim b As Double
        Dim x As Double
        Dim y As Double
        a = CDbl(TxtA.Text)
        b = CDbl(TxtB.Text)
        x = CDbl(TxtX.Text)
        
        y = a + ((x - 24) * b)
         Label5.Caption = y
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Label5.Caption = ""
End Sub
