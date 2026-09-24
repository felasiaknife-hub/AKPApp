VERSION 5.00
Begin VB.Form frAlertShow 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEnd 
      Caption         =   "หยุดเตือนข้อความนี้"
      Height          =   345
      Left            =   240
      TabIndex        =   4
      Top             =   3330
      Width           =   1725
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   3330
      Width           =   1245
   End
   Begin VB.Image imgAlert 
      Height          =   495
      Left            =   210
      Picture         =   "frAlertShow.frx":0000
      Stretch         =   -1  'True
      Top             =   150
      Width           =   555
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -810
      Top             =   3240
      Width           =   9825
   End
   Begin VB.Label lbCond1 
      Caption         =   "คำเตือนนี้จะแสดงทุกครั้งที่เปิดและปิดโปรแกรม"
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   2730
      Width           =   5520
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   1665
      Left            =   240
      Top             =   990
      Width           =   5535
   End
   Begin VB.Label lbMessage 
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   1395
      Left            =   360
      TabIndex        =   2
      Top             =   1110
      Width           =   5265
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำเตือน........"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   840
      TabIndex        =   0
      Top             =   90
      Width           =   1620
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   0
      Picture         =   "frAlertShow.frx":0209
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10050
   End
End
Attribute VB_Name = "frAlertShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curAlert_ As AWSAlert

Private Sub cmdClose_Click()
    If chkEnd.Value = 1 Then
        DBManager.GetDBConnection.Execute "Update tbAlert Set isActive=0 Where AlertID='" & curAlert_.AlertID & "'"
    End If
    Unload Me
End Sub

Public Sub setAlertItem(alertItem As AWSAlert)
Dim remStr1 As String
Dim remStr2 As String
    With alertItem
        lbMessage.Caption = .Message
        Label1.Caption = "คำเตือน...... " & FormatYMD_to_DMY(.startDate, "/", "/")
        If .AlertAtEnd = 0 And .AlertAtStart = 1 Then
            remStr1 = "คำเตือนนี้จะแสดงทุกครั้งที่เปิดโปรแกรม"
        ElseIf .AlertAtEnd = 1 And .AlertAtStart = 1 Then
            remStr1 = "คำเตือนนี้จะแสดงทุกครั้งที่เปิดและก่อนปิดโปรแกรม"
        ElseIf .AlertAtEnd = 1 And .AlertAtStart = 0 Then
            remStr1 = "คำเตือนนี้จะแสดงทุกครั้งก่อนปิดโปรแกรม"
        End If
        If .ManualStop = 1 Then
            remStr2 = "จะสิ้นสุดการเตือนเมื่อผู้ใช้สั่งหยุด"
        Else
            remStr2 = "จะแสดงข้อความเตือนจนถึงวันที่ " & FormatYMD_to_DMY(.endDate, "/", "/")
        End If
        If Trim(remStr1) <> "" Then
            lbCond1.Caption = remStr1 & vbLf & remStr2
        Else
            lbCond1.Caption = remStr2
        End If
    End With
    Set curAlert_ = alertItem
End Sub

Private Sub Form_Load()
    'If Dir(IconPath & "\alert2.ico") <> "" Then imgAlert.Picture = LoadPicture(IconPath & "\alert2.ico")

End Sub
