VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2625
   ClientLeft      =   15
   ClientTop       =   -75
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frLogin.frx":0ECA
   ScaleHeight     =   2625
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboConnType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frLogin.frx":324A8A
      Left            =   975
      List            =   "frLogin.frx":324A8C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1260
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "เข้าระบบ"
      Default         =   -1  'True
      Height          =   795
      Left            =   2130
      Picture         =   "frLogin.frx":324A8E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1590
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1335
      TabIndex        =   0
      Top             =   570
      Width           =   2625
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1335
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   930
      Width           =   2625
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   3300
      Picture         =   "frLogin.frx":324B77
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1590
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   90
      TabIndex        =   8
      Top             =   1305
      Width           =   4755
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Secure Login System..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   6
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   5
      Top             =   960
      Width           =   690
   End
   Begin VB.Image Image2 
      Height          =   2955
      Left            =   -60
      Picture         =   "frLogin.frx":3250F3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5190
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

Private Sub CommandClosed()
    EndApplication = True
'    Unload frmClient
'    Localization.SetDefaultRegionalSettings DefaultLCID
    Unload Me
End Sub

Private Sub CommandLogin()
Dim isValid As Boolean
Dim mySysFile As New FileSystemObject
Dim myTxtStream As TextStream
Dim myPath As String
Dim myEnc As New RC4Crypt
Dim M, D As Integer
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
    
    SaveSetting App.EXEName, Me.Name, cboConnType.Name, cboConnType.ListIndex
    ConnectServer = cboConnType.ListIndex
    If GetSystemConfig = False Then Exit Sub

    Set CurrentUser = ProcessLogin(Trim(txtUserName.Text), Trim(txtPassword.Text))
    If Not CurrentUser Is Nothing Then
        Unload Me
        EndApplication = False
        
        '*** Check Application Version
        Dim i As Integer, CheckVersion As Boolean
        Dim UserGrp As New AWSUserGroup
        Dim tSql As String, rs As New ADODB.Recordset
        CheckVersion = SkipCheckVersion
        If Not CurrentUser Is Nothing Then
            For i = 1 To CurrentUser.getUserGroups.Count
                Set UserGrp = CurrentUser.getUserGroups(i)
                If Trim(UserGrp.Name) = "Administrator" Or Trim(UserGrp.Name) = "IT" Then
                    If SkipCheckVersion <> CheckVersion Then SkipCheckVersion = CheckVersion
                    Exit For
                Else
                    SkipCheckVersion = False
                End If
            Next i
        End If
    
        If CurrentUser Is Nothing Then Exit Sub
'        If SkipCheckVersion = False Then
'            tSql = "SELECT ApplicationName, vApp, AppMsg FROM ApplicationTB WHERE ApplicationName = '" & App.EXEName & "' OR ApplicationName = '" & App.ProductName & "'"
'            Set rs = GetRS(tSql, adOpenStatic, adLockOptimistic)
'            If rs.RecordCount > 0 Then
''                If AppNewVersion(App.Path & "\" & App.EXEName) Then
'                If Trim("" & rs!AppMsg) <> "" Then
'                    MsgBox rs!AppMsg, vbInformation
'                    End '*** End Application
'                End If
'                If AppVersion < rs!vApp Then
'                    MsgBox "โปรแกรม Version ไม่ตรงกับ Server กรุณาแจ้งฝ่าย IT เพื่ออัพเดทโปรแกรม" & vbCrLf & _
'                    "App Version " & AppVersion & " = Server Version " & rs!vApp
'                    End '*** End Application
'                End If
'            Else
'                MsgBox "โปรแกรมที่ท่านใช้เก่าเกินไป กรุณาแจ้งฝ่าย IT เพื่ออัพเดทโปรแกรม" & vbCrLf & _
'                "Application Name " & App.EXEName & " " & AppVersion
'                End '*** End Application
'            End If
'        End If
        '*** End check version
    
        frmMain.InitializeMenu
        frmMain.ActivateForm
        frmMain.Show
    Else
        errCount_ = errCount_ + 1
        If errCount_ < 3 Then
            lblStatus.Caption = "ชื่อหรือรหัสผ่านไม่ถูกต้อง ตรวจสอบชื่อ รหัสผ่าน และลองใหม่อีกครั้ง" & " Retry Login: " & 3 - errCount_
            txtUserName.SetFocus
            EndApplication = False
            Exit Sub
        Else
            lblStatus.Caption = "ชื่อหรือรหัสผ่านไม่ถูกต้อง ตรวจสอบชื่อ รหัสผ่าน และลองใหม่อีกครั้ง" & " Retry Login: " & 3 - errCount_
            MsgBox "ข้อมูลการเข้าระบบไม่ถูกต้องติดต่อกันครบ 3 ครั้ง โปรแกรมไม่สามารถทำงานต่อได้", vbExclamation
            EndApplication = True
            Unload Me
        End If
    End If
End Sub

Private Sub cboConnType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CommandLogin
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not TypeOf Me.ActiveControl Is CommandButton Then
            CommandClosed
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim mySysFile As New FileSystemObject
Dim myPath As String
Dim myEnc As New RC4Crypt
Dim myTxtStream As TextStream
Dim tmpStr As String
Dim M, D As Integer
    KeyPreview = True
    SetAlwaysOnTopMode Me, True
    MakeFormRounded Me, 25, , 2
    
    With cboConnType
         .AddItem "Local Network 1"
        .AddItem "Internet Network SPK"
        .AddItem "Internet Network BKK"
        .ListIndex = 1
    End With
    
'    cboConnType.ListIndex = IIf(GetSetting(App.EXEName, Me.Name, cboConnType.Name, cboConnType.ListIndex) = Null, 0, _
'                                                                GetSetting(App.EXEName, Me.Name, cboConnType.Name, cboConnType.ListIndex))
    
    If Right(App.Path, 1) = "\" Then
        myPath = App.Path
    Else
        myPath = App.Path & "\"
    End If
    
'    m = Month(Now)
'    d = Day(Now)
'    If (m = 1 And d = 1) Or (m = 5 And d = 5) Or (m = 9 And d = 9) Or (m = 11 And d = 11) Then
'        modUtil.getPassword Replace(modUtil.GetVolumeID + Year(Now) + Month(Now), "-", "")
'    Else
'        modUtil.getPassword modUtil.GetVolumeID
'    End If

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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
'    Me.Height = 2940
'    Me.Width = 4590
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CommandLogin
    End If
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub

Private Sub ShowID1ID2(isShow As Boolean)
    If isShow Then
        Me.Height = 3390
        cmdOK.Top = 2160
        cmdClose.Top = 2160
        Label1(2).Visible = True
    Else
        Me.Height = 2730
        cmdOK.Top = 1530
        cmdClose.Top = 1530
        Label1(2).Visible = False
    End If
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim EncObj As New RC4Crypt
Dim sqlStr As String
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    ElseIf KeyCode = 115 Then
        If Trim(txtUserName.Text) <> "" Then
            sqlStr = "SELECT * FROM tbCompanyStaff WHERE LoginName='" & Trim(txtUserName.Text) & "'"
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
            If tmpRec.EOF = False Then
                MsgBox EncObj.DecryptString(Trim("" & tmpRec("LoginPwd")), ENC_KEY)
            End If
        End If
    End If
    Set tmpRec = Nothing
    Set EncObj = Nothing
    Exit Sub
ErrD:
    MsgBox "", vbCritical
    Exit Sub
End Sub



