VERSION 5.00
Begin VB.Form frmChangeTripNo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "เปลี่ยนเลขที่ใบคุม F153"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOldPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2130
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   390
      Width           =   2025
   End
   Begin VB.TextBox txtNewPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2130
      TabIndex        =   1
      Top             =   810
      Width           =   2025
   End
   Begin VB.TextBox txtConPwd 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2130
      TabIndex        =   2
      Top             =   1230
      Width           =   2025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   795
      Left            =   2610
      Picture         =   "frChangeTripNo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   1065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   3720
      Picture         =   "frChangeTripNo.frx":00E9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบคุมเดิม"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1050
      TabIndex        =   7
      Top             =   450
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบคุมใหม่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1020
      TabIndex        =   6
      Top             =   870
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยืนยันเลขที่ใบคุมใหม่"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   630
      TabIndex        =   5
      Top             =   1260
      Width           =   1440
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   0
      Picture         =   "frChangeTripNo.frx":0665
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4890
   End
End
Attribute VB_Name = "frmChangeTripNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CallerForm As Form
Dim tbJobDataTimeTable As String

Public Sub JobDataTimeTable(JobData As String)
    tbJobDataTimeTable = JobData
End Sub

Public Sub SetCallerForm(frm As Form)
    Set CallerForm = frm
End Sub

Public Sub SetTripNo(selNo As String)
    txtOldPwd.Text = Trim(selNo)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tNo1$, tNo2$, sqlStr As String, Index As Long
    If tbJobDataTimeTable <> "tbJobDataTimeTableBox" Then
        tbJobDataTimeTable = "tbJobDataTimeTable"
        Index = 8
    Else
        Index = 1
    End If
    If Trim(txtNewPwd.Text) = "" Then
        MsgBox "กรุณาระบุเลขที่ใบคุมใหม่", vbExclamation
        txtNewPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtConPwd.Text) = "" Then
        MsgBox "กรุณายืนยันเลขที่ใบคุมใหม่", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtNewPwd.Text) <> Trim(txtConPwd.Text) Then
        MsgBox "การยืนยันเลขที่ใบคุมใหม่ไม่ถูกต้อง การยืนยันและเลขที่ใบคุมใหม่จะต้องตรงกัน", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
    If InStr(1, txtNewPwd.Text, "-") = 0 Then
        sqlStr = "SELECT * FROM " & tbJobDataTimeTable & " WHERE TimeTableNo='" & Trim(txtNewPwd.Text) & "' And ISNULL(isCancel, 'N') ='N'"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        If tmpRec.EOF Then
            MsgBox "ไม่พบเลขที่ใบคุมใหม่ในฐานข้อมูล กรุณาตรวจสอบความถูกต้อง", vbExclamation
            txtNewPwd.SetFocus
            tmpRec.Close
            Set tmpRec = Nothing
            Exit Sub
        End If
        tmpRec.Close
    Else
        tNo1 = Trim(Split(txtNewPwd.Text, "-")(0))
        tNo2 = Left(Trim(tNo1), 4) & Trim(Split(txtNewPwd.Text, "-")(1))
        sqlStr = "SELECT * FROM  " & tbJobDataTimeTable & "  WHERE (TimeTableNo='" & tNo1 & "' And ISNULL(isCancel, 'N') ='N')"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        If tmpRec.EOF Then
            MsgBox "ไม่พบเลขที่ใบคุม " & tNo1 & " ในฐานข้อมูล กรุณาตรวจสอบความถูกต้อง", vbExclamation
            txtNewPwd.SetFocus
            tmpRec.Close
            Set tmpRec = Nothing
            Exit Sub
        End If
        tmpRec.Close
        sqlStr = "SELECT * FROM  " & tbJobDataTimeTable & "  WHERE (TimeTableNo='" & tNo2 & "' And ISNULL(isCancel, 'N') ='N')"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        If tmpRec.EOF Then
            MsgBox "ไม่พบเลขที่ใบคุม " & tNo2 & " ในฐานข้อมูล กรุณาตรวจสอบความถูกต้อง", vbExclamation
            txtNewPwd.SetFocus
            tmpRec.Close
            Set tmpRec = Nothing
            Exit Sub
        End If
        tmpRec.Close
    End If
    Set tmpRec = Nothing
    If MsgBox("ท่านต้องการเปลี่ยนเลขที่ใบคุมหรือไม่?", vbQuestion + vbYesNo) = vbYes Then

        DBConnExc "Update  " & tbJobDataTimeTable & "  Set TimeTableNo= '" & Trim(txtNewPwd.Text) & "', isMain='N' WHERE TimeTableNo='" & Trim(txtOldPwd.Text) & "'"
    
        CallerForm.UpdateTimeTableNo Trim(txtNewPwd.Text), Index
        MsgBox "เลขที่ใบคุมถูกเปลี่ยนเรียบร้อยแล้ว", vbExclamation
        Unload Me
    End If
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

