VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "นำเข้าข้อมูล"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   4470
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4980
      TabIndex        =   9
      Text            =   "*.*"
      Top             =   120
      Width           =   2475
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   405
      Left            =   5790
      TabIndex        =   4
      Top             =   4020
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   405
      Left            =   4080
      TabIndex        =   3
      Top             =   4020
      Width           =   1665
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   4050
      TabIndex        =   2
      Top             =   750
      Width           =   3405
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3690
      Left            =   120
      TabIndex        =   1
      Top             =   750
      Width           =   3765
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "รูปแบบไฟล์"
      Height          =   195
      Index           =   3
      Left            =   4080
      TabIndex        =   8
      Top             =   150
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "เลือกไฟล์"
      Height          =   195
      Index           =   2
      Left            =   4080
      TabIndex        =   7
      Top             =   540
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "เลือกโฟลเดอร์"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   540
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "เลือกไดรฟ์"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim defaultPath_ As String
Dim callerForm As Form

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strPath As String
    If Not callerForm Is Nothing Then
        strPath = File1.Path
        If Right(strPath, 1) <> "\" Then
            strPath = strPath & "\"
        End If
        'Me.Height = 5205
        'pb1.Visible = True
        If Trim(File1.FileName) = "" Then
            MsgBox "กรุณาเลือกไฟล์ข้อมูล", vbExclamation
            Exit Sub
        End If
        callerForm.KickOfProcess strPath & File1.FileName
        'Me.Height = 4860
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveErr:
    Dir1.Path = Drive1.Drive
    Exit Sub
DriveErr:
    MsgBox "ไดรฟ์ยังไม่พร้อมใช้งานหรือขัดข้อง กรุณาตรวจสอบ", vbExclamation
    Drive1.Drive = Dir1.Path
    Exit Sub
End Sub

Public Property Get getCurrentPath() As String
    getCurrentPath = Trim(Dir1.Path)
End Property

Public Property Get getFileName() As String
    getFileName = Trim(File1.FileName)
End Property

Public Property Get DefaultPath() As String
    DefaultPath = Trim(defaultPath_)
End Property
Public Property Let DefaultPath(ByVal newVal As String)
On Error GoTo pathErr:
    If Trim(defaultPath_) <> Trim(newVal) Then
        If Trim(Dir(newVal, vbDirectory)) <> "" Then
            defaultPath_ = Trim(newVal)
            Dir1.Path = defaultPath_
            Drive1.Drive = Dir1.Path
        End If
    End If
    Exit Property
pathErr:
    Dir1.Path = Drive1.Drive
    defaultPath_ = Dir1.Path
    Exit Property
End Property

Public Property Get FilePattern() As String
    FilePattern = File1.Pattern
End Property
Public Property Let FilePattern(ByVal newVal As String)
    File1.Pattern = newVal
    Text1.Text = newVal
    File1.Refresh
End Property

Public Sub setCallerForm(frm As Form)
    Set callerForm = frm
End Sub

Private Sub File1_DblClick()
    If File1.ListCount > 0 Then
        cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        File1.Pattern = Text1.Text
    End If
End Sub
