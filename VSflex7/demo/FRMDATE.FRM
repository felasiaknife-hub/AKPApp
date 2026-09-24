VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date Picker"
   ClientHeight    =   885
   ClientLeft      =   4545
   ClientTop       =   2460
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   1104
   End
   Begin VB.ComboBox txtMonth 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmdate.frx":0000
      Left            =   1230
      List            =   "Frmdate.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   1104
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   540
      Width           =   1104
   End
   Begin VB.Label lblCol 
      Caption         =   "Label4"
      Height          =   300
      Left            =   3384
      TabIndex        =   8
      Top             =   -144
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Label lblRow 
      Caption         =   "Label4"
      Height          =   300
      Left            =   3276
      TabIndex        =   7
      Top             =   -144
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   3450
   End
   Begin VB.Label Label3 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1260
      TabIndex        =   2
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2430
      TabIndex        =   4
      Top             =   360
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   600
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dt As Variant

Sub UpdateDate()
    If Not Enabled Then Exit Sub
    If Not IsDate(txtMonth & "/" & txtDay & "/" & txtYear) Then Beep: Exit Sub
    dt = txtMonth & "/" & txtDay & "/" & txtYear
    lblDate = Format(dt, "Long Date")
    Tag = dt
End Sub

Private Sub Form_Activate()
    dt = Tag
    Enabled = False
    lblDate = Format(dt, "Long Date")
    txtYear = Year(dt)
    txtDay = Day(dt)
    txtMonth.ListIndex = Month(dt) - 1
    Enabled = True
End Sub

Private Sub Form_Deactivate()
    
    On Error Resume Next
    
    ' update grid value
    If IsDate(Tag) Then frmMenu!fa.Cell(flexcpText, lblRow, lblCol) = Tag
    
    ' go away
    Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    ' esc clears, then quits
    If KeyAscii = 27 Then Tag = "": Hide
    
    ' enter quits
    If KeyAscii = 13 Then Hide

End Sub

Private Sub txtDay_Change()
    UpdateDate
End Sub


Private Sub txtDay_GotFocus()
    txtDay.SelStart = 0
    txtDay.SelLength = 30000
    txtDay.MaxLength = 2
End Sub


Private Sub txtMonth_Change()
    UpdateDate
End Sub

Private Sub txtMonth_Click()
    UpdateDate
End Sub

Private Sub txtYear_Change()
    UpdateDate
End Sub


Private Sub txtYear_GotFocus()
    txtYear.SelStart = 0
    txtYear.SelLength = 30000
    txtYear.MaxLength = 4
End Sub


