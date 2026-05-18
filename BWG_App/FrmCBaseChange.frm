VERSION 5.00
Begin VB.Form FrmCBaseChange 
   Caption         =   "เปลี่ยนฐานลูกค้า"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtBase 
      Height          =   285
      Left            =   2295
      TabIndex        =   5
      Top             =   540
      Width           =   1635
   End
   Begin VB.TextBox TxtCode 
      Height          =   285
      Left            =   2295
      TabIndex        =   4
      Top             =   180
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ปิด"
      Height          =   465
      Left            =   2835
      TabIndex        =   3
      Top             =   1350
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "บันทึก"
      Height          =   465
      Left            =   1665
      TabIndex        =   2
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "รหัสที่ได้ : "
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   900
      Width           =   720
   End
   Begin VB.Label LblID 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2295
      TabIndex        =   7
      Top             =   900
      Width           =   45
   End
   Begin VB.Label LblNum 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   990
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "เปลี่ยนเป็นฐาน : "
      Height          =   195
      Left            =   1035
      TabIndex        =   1
      Top             =   585
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "รหัสลูกค้าที่ต้องการเปลี่ยน : "
      Height          =   195
      Left            =   315
      TabIndex        =   0
      Top             =   225
      Width           =   1935
   End
End
Attribute VB_Name = "FrmCBaseChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSave_Click()
    If UCase(Left(TxtCode.Text, 2)) = "MS" Then
        LblNum.Caption = Mid(Trim(TxtCode.Text), 3, Len(Trim(TxtCode.Text)))
    Else
        LblNum.Caption = Mid(Trim(TxtCode.Text), 2, Len(Trim(TxtCode.Text)))
    End If
    LblID.Caption = Trim(UCase(TxtBase.Text)) & Trim(LblNum.Caption)
    Dim STR As String
    DBConnExc "Update tbcustomer Set CustomerCode='" & Trim(LblID.Caption) & "',CustTypeFlag='" & Trim(UCase(TxtBase.Text)) & "',changedate='" & FormatYMD_Thai(Day(Date), Month(Date), Year(Date), "/") & "',LastChangeUser='" & CurrentUser.StaffFName + " " & CurrentUser.StaffLName & "' WHERE CustomerCode='" & Trim(TxtCode.Text) & "'"
        MsgBox "แก้ไขข้อมูลเรียบร้อย", vbInformation + vbOKOnly, "เสร็จสิ้นการแก้ไข"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub TxtBase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBase.SetFocus
    End If
End Sub
