VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "เข้าระบบ"
      Default         =   -1  'True
      Height          =   795
      Left            =   2130
      Picture         =   "frLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1590
      Width           =   1125
   End
   Begin VB.TextBox txtID2 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   450
      TabIndex        =   3
      Top             =   1980
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox txtID1 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1590
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   0
      Top             =   390
      Width           =   2025
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   810
      Width           =   2025
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   3300
      Picture         =   "frLogin.frx":00E9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID2"
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
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   2010
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID1"
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
      Left            =   60
      TabIndex        =   8
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อที่ใช้เข้าระบบ"
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
      Left            =   360
      TabIndex        =   7
      Top             =   420
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผ่าน"
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
      Left            =   870
      TabIndex        =   6
      Top             =   840
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   -30
      X2              =   4650
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Image Image2 
      Height          =   3075
      Left            =   -420
      Picture         =   "frLogin.frx":0665
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4890
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim errCount_ As Integer
Dim colCompany_ As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim isValid As Boolean
Dim mySysFile As New FileSystemObject
Dim myTxtStream As TextStream
Dim myPath As String
Dim myEnc As New RC4Crypt
Dim m, d As Integer
    If Trim(txtUserName.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อที่ใช้ในการเข้าระบบ", vbExclamation
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "กรุณาใส่รหัสผ่าน", vbExclamation
        txtPassword.SetFocus
        Exit Sub
    End If
    If txtID2.Visible Then
        If Trim(txtID2.Text) <> Trim(modUtil.ProductPassword) Then
            MsgBox "ID 2 ไม่ถูกต้อง กรุณาติดต่อ โทร. 06-677-7237", vbExclamation
            txtID2.SetFocus
            Exit Sub
        Else
            m = Month(Now)
            d = Day(Now)
            If Not ((m = 1 And d = 1) Or (m = 5 And d = 5) Or (m = 9 And d = 9) Or (m = 11 And d = 11)) Then
                If Right(App.Path, 1) = "\" Then
                    myPath = App.Path
                Else
                    myPath = App.Path & "\"
                End If
                Set myTxtStream = mySysFile.OpenTextFile(myPath & "msfix.txk", ForWriting, True)
                myTxtStream.WriteLine myEnc.EncryptString(ProductPassword, modMain.ENC_KEY)
                myTxtStream.Close
                Set myTxtStream = Nothing
                Set mySysFile = Nothing
                Set myEnc = Nothing
            End If
        End If
    End If
    Set CurrentUser = ProcessLogin(Trim(txtUserName.Text), Trim(txtPassword.Text))
    If Not CurrentUser Is Nothing Then
        Unload Me
        'MsgBox "Login Succeed"
'        frmMDIMain.ActiveForm
        frmMDIMain.Show
'        EnableCheckUpdate
    Else
        errCount_ = errCount_ + 1
        If errCount_ < 3 Then
            MsgBox "ข้อมูลการเข้าระบบไม่ถูกต้อง กรุณาตรวจสอบชื่อในการเข้าระบบ, รหัสผ่าน และบริษัทที่ท่านสังกัด", vbExclamation
            txtUserName.SetFocus
            Exit Sub
        Else
            MsgBox "ข้อมูลการเข้าระบบไม่ถูกต้องติดต่อกันครบ 3 ครั้ง โปรแกรมไม่สามารถทำงานต่อได้", vbExclamation
            Unload Me
        End If
    End If
End Sub


Private Sub Form_Load()
Dim mySysFile As New FileSystemObject
Dim myPath As String
Dim myEnc As New RC4Crypt
Dim myTxtStream As TextStream
Dim tmpStr As String
Dim m, d As Integer
    
    If Right(App.Path, 1) = "\" Then
        myPath = App.Path
    Else
        myPath = App.Path & "\"
    End If
    m = Month(Now)
    d = Day(Now)
    If (m = 1 And d = 1) Or (m = 5 And d = 5) Or (m = 9 And d = 9) Or (m = 11 And d = 11) Then
        modUtil.getPassword Replace(modUtil.GetVolumeID + Year(Now) + Month(Now), "-", "")
    Else
        modUtil.getPassword modUtil.GetVolumeID
    End If

'    ShowID1ID2 False
'    If mySysFile.FileExists(myPath & "msfix.txk") = False Then
'        ShowID1ID2 True
'        txtID1.Text = Trim(modUtil.ProductID)
'    Else
'        txtID1.Text = Trim(modUtil.ProductID)
'        Set myTxtStream = mySysFile.OpenTextFile(myPath & "msfix.txk", ForReading, True)
'        tmpStr = myEnc.DecryptString((myTxtStream.ReadLine), modMain.ENC_KEY)
'        If Trim(tmpStr) = Trim(modUtil.ProductPassword) Then
'            ShowID1ID2 False
'        Else
'            ShowID1ID2 True
'        End If
'        myTxtStream.Close
'        Set myTxtStream = Nothing
'        Set myEnc = Nothing
'    End If
'    Set mySysFile = Nothing
End Sub


Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub

Private Sub ShowID1ID2(isShow As Boolean)
    If isShow Then
        Me.Height = 3390
        Line1.Y1 = 2070
        Line1.Y2 = 2070
        cmdOK.Top = 2160
        cmdClose.Top = 2160
        Label1(2).Visible = True
        txtID1.Visible = True
        Label1(3).Visible = True
        txtID2.Visible = True
    Else
        Me.Height = 2730
        Line1.Y1 = 1440
        Line1.Y2 = 1440
        cmdOK.Top = 1530
        cmdClose.Top = 1530
        Label1(2).Visible = False
        txtID1.Visible = False
        Label1(3).Visible = False
        txtID2.Visible = False
    End If
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim encObj As New RC4Crypt
    If KeyCode = 115 Then
        If Trim(txtUserName.Text) <> "" Then
            Set tmpRec = GetRS("Select * from tbCompanyStaff Where LoginName='" & Trim(txtUserName.Text) & "'", adOpenForwardOnly, adLockReadOnly)
'            tmpRec.Open "Select * from tbCompanyStaff Where LoginName='" & Trim(txtUserName.Text) & "'", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
            If tmpRec.EOF = False Then
                MsgBox encObj.DecryptString(Trim("" & tmpRec("LoginPwd")), ENC_KEY)
            End If
        End If
    ElseIf KeyCode = 116 Then
        If Trim(txtUserName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
            DBManager.GetDBConnection.Execute "Update tbTeacherData Set LoginPwd='" & encObj.EncryptString(Trim("" & txtPassword.Text), ENC_KEY) & "' Where LoginName='" & Trim(txtUserName.Text) & "'"
        End If
    End If
    Set tmpRec = Nothing
    Set encObj = Nothing
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

