VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServer 
   BackColor       =   &H00800000&
   Caption         =   "   Dual-Gate Advanced Server"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13215
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerEMail 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5400
      Top             =   7800
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   7560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "frmServer.frx":11C2
      Top             =   195
      Width           =   5535
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   720
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1200
      Top             =   7920
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame FrameSocket 
      Caption         =   "           Sever Properties                                                Clients Connected         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txtEMail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         HideSelection   =   0   'False
         Left            =   1440
         TabIndex        =   8
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txtMaxClients 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         HideSelection   =   0   'False
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         HideSelection   =   0   'False
         Left            =   840
         MaxLength       =   5
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         HideSelection   =   0   'False
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
      Begin VB.ListBox lstClients 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   5040
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Startup EMail:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clients:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MouseIcon       =   "frmServer.frx":1249
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Frame FrameDB 
      Caption         =   "Connected Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4455
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   7095
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   37
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   28
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtPwd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   36
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   2
         Left            =   1800
         TabIndex        =   34
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtPwd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   1
         Left            =   1800
         TabIndex        =   25
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtPwd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         ItemData        =   "frmServer.frx":1553
         Left            =   360
         List            =   "frmServer.frx":155D
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3240
         Width           =   1335
      End
      Begin VB.ComboBox cboDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "frmServer.frx":1578
         Left            =   360
         List            =   "frmServer.frx":1582
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cboDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "frmServer.frx":159D
         Left            =   360
         List            =   "frmServer.frx":15A7
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDB 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6600
         TabIndex        =   32
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdDB 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6600
         TabIndex        =   23
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdDB 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "For SQL Server databases use \\server_name\database name"
         Top             =   3240
         Width           =   4695
      End
      Begin VB.TextBox txtDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "For SQL Server databases use \\server_name\database name"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.TextBox txtDB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "For SQL Server databases use \\server_name\database name"
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PWD:"
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
         Left            =   3480
         TabIndex        =   35
         Top             =   3960
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UID:"
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
         Left            =   1320
         TabIndex        =   33
         Top             =   3960
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PWD:"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   2520
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UID:"
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
         Left            =   1320
         TabIndex        =   24
         Top             =   2520
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PWD:"
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
         Left            =   3480
         TabIndex        =   17
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UID:"
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
         Left            =   1320
         TabIndex        =   15
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3."
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
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2."
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
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1."
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
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   150
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuChangeClients 
         Caption         =   "Change Clients   "
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Password"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFolderName As String
Dim lngFreeFileNo As Long
Dim strHeader As String
Dim strData As String
Dim strFileName As String

Dim strSeverPwd As String
Dim strPwd
Dim boolConfirm
Dim intX As Integer
Public intMax As Integer

Private Sub cboDB_Change(Index As Integer)
If UCase(cboDB(Index).Text) = "MS ACCESS" Then
    txtDB(Index).Locked = True
    txtDB(Index).Text = ""
    cmdDB(Index).Enabled = True
    txtUID(Index).Locked = True
Else
    txtDB(Index).Locked = False
    txtDB(Index).Text = ""
    cmdDB(Index).Enabled = False
    txtUID(Index).Locked = False
End If
End Sub

Private Sub cboDB_Click(Index As Integer)
Call cboDB_Change(Index)
End Sub

Private Sub cmdClear_Click(Index As Integer)
txtDB(Index).Text = ""
txtUID(Index).Text = ""
txtPwd(Index).Text = ""
End Sub

Private Sub cmdDB_Click(Index As Integer)
CommonDialog1.Filter = "MS Access Database(*.mdb)|*.mdb"
On Error GoTo err:
CommonDialog1.ShowOpen

txtDB(Index).Text = CommonDialog1.FileName
Exit Sub

err:
End Sub

Private Sub cmdStart_Click()
If IsNumeric(txtPort.Text) = False Then
    GoTo err
End If

If cmdStart.Caption = "&Start Server" Then
    If chkDatabases = True Then
        On Error GoTo err1
        sckServer(0).LocalPort = txtPort.Text
        sckServer(0).Listen
        txtServer.Text = UCase(sckServer(0).LocalHostName) & " (" & sckServer(0).LocalIP & ")"
        cmdStart.Caption = "&Stop Server"
        Call AddSeverMsg("Server Initialized on Port " & txtPort.Text & "...")
        Call saveDataToReg
        
        FrameDB.Enabled = False
        FrameSocket.Enabled = False
        mnuOptions.Enabled = False
        
        txtMsg.Text = Replace(txtMsg.Text, "Server Status : Terminated", "Server Status : ")
        txtMsg.Text = Replace(txtMsg.Text, "Server Status : ", "Server Status : Running...")
        Me.Caption = " Dual-Gate Advanced Server : Running..."
        frmServer.txtMsg.SelStart = Len(frmServer.txtMsg.Text)
        
        timerEMail.Enabled = True
    End If
Else
    Call Form_Unload(0)
    Me.Caption = " Dual-Gate Advanced Server : Stopped"
End If
Exit Sub

err:
Call AddSeverMsg("ERROR : Set your server port to continue.")
On Error Resume Next
txtPort.SetFocus
Exit Sub

err1:
Call AddSeverMsg("ERROR : Port number(address) already in use")
Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo err
Set FSO = CreateObject("Scripting.FileSystemObject")
Call chkFolderPath(App.Path & "\Files\", True)

If App.PrevInstance = True Then
    MsgBox "Another instance of Dual-Gate Advanced Server is already running.", vbExclamation, "Dual-Gate Server"
    Unload Me
    Exit Sub
End If
Call changeDateFormat

txtServer.Text = UCase(sckServer(0).LocalHostName) & " (" & sckServer(0).LocalIP & ")"
cboDB(0).Text = cboDB(0).List(0)
cboDB(1).Text = cboDB(0).List(0)
cboDB(2).Text = cboDB(0).List(0)

Call loadFromReg
Call cmdStart_Click

Dim cl As New cAppHider
cl.HideApplication
Exit Sub

err:
Call AddSeverMsg(err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If cmdStart.Caption <> "&Start Server" Then
    boolConfirm = MsgBox("Are you sure you want to terminate the server ?" & vbCrLf & vbCrLf & "This will disconnect all clients.", vbYesNoCancel + vbExclamation + vbDefaultButton3, "Dual-Gate Server")
    If boolConfirm = vbYes Then
        If Len(strSeverPwd) > 0 Then
            strPwd = InputBox("Enter your server password.", "Enter password")
        Else
            strPwd = strSeverPwd
        End If
        If strPwd = strSeverPwd Then
            intMax = 1
            sckServer(0).Close
            While intMax <= sckServer.UBound
                Unload sckServer(intMax)
                sckServer(intMax).Close
                intMax = intMax + 1
            Wend
            cmdStart.Caption = "&Start Server"
            txtMsg.Text = Replace(txtMsg.Text, "Server Status : Running...", "Server Status : ")
            txtMsg.Text = Replace(txtMsg.Text, "Server Status : ", "Server Status : Terminated")
            FrameDB.Enabled = True
            FrameSocket.Enabled = True
            mnuOptions.Enabled = True
            lstClients.Clear
            
            Call AddSeverMsg("Server Terminated..." & vbCrLf)
        Else
            MsgBox "Invalid Password !!!", vbExclamation, "Dual-Gate Server"
            Cancel = 1
        End If
    Else
        Cancel = 1
    End If
End If
End Sub

Private Sub mnuChangeClients_Click()
strPwd = InputBox("Enter your password.", "Configure Clients...")
If strPwd = Replace(Format(Now, "ddmmyyhhmmAMPM"), Format(Now, "AMPM"), "") Then
    strPwd = InputBox("Edit No. of Clients.", "Configure Clients...")
    If IsNumeric(strPwd) = True Then
        txtMaxClients.Text = Format(strPwd, "###0")
        txtMaxClients.Text = Replace(txtMaxClients.Text, "-", "")
        Call AddSeverMsg("Configuration of clients successful.")
    Else
        MsgBox "Enter valid number of clients !!!", vbExclamation, "Dual-Gate Server"
    End If
Else
    Call AddSeverMsg("Incorrect password entered.")
    MsgBox "Invalid Password !!!", vbExclamation, "Dual-Gate Server"
End If
End Sub

Private Sub mnuPassword_Click()
strPwd = InputBox("Enter your password.", "Enter Password...")
If strPwd = strSeverPwd Then
    strPwd = InputBox("Enter your new server password.", "Edit Password...")
    strSeverPwd = strPwd
    SaveSetting App.EXEName, "Config", "PWD", Encrypt(strSeverPwd, strPassword)
    Call AddSeverMsg("Server password modified successfully.")
    MsgBox "Password changed successfully.", vbInformation, "Dual-Gate Server"
Else
    Call AddSeverMsg("Incorrect password entered.")
    MsgBox "Invalid Password !!!", vbExclamation, "Dual-Gate Server"
End If
End Sub

Private Sub sckServer_Close(Index As Integer)
Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " Disconnected")
lstClients.RemoveItem Index - 1
Unload sckServer(Index)
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim intClient As Integer
intClient = 0
If Index = 0 Then
    Call AddSeverMsg(sckServer(Index).RemoteHostIP & " requesting connection...")
    ' this is to prevent duplicate connection from same client
    While intClient <= lstClients.ListCount
        If sckServer(Index).RemoteHostIP = lstClients.List(intClient) Then
            Call AddSeverMsg("ERROR : The client (" & sckServer(Index).RemoteHostIP & ") is already connected")
            Exit Sub
        End If
        intClient = intClient + 1
    Wend
    'ends code
    
    'checks the allowed maximum no of clients
    If lstClients.ListCount + 1 <= CInt(txtMaxClients.Text) Then
        Load sckServer(sckServer.UBound + 1)
        sckServer(sckServer.UBound).LocalPort = 0
        sckServer(sckServer.UBound).Accept requestID
        Call AddSeverMsg(sckServer(sckServer.UBound).RemoteHostIP & " connected.")
        lstClients.AddItem sckServer(sckServer.UBound).RemoteHostIP
        
        'Client authentication
        Call AddSeverMsg("Authenticating client (" & sckServer(Index).RemoteHostIP & ")")
        sckServer(sckServer.UBound).sendData "#^PASSWORD^#"
        boolPassword = False
        Call Wait(0.1)
        If boolPassword = False Then
            Call sckServer_Close(sckServer.UBound)
            Call AddSeverMsg("ERROR : Authentication failed for client. Connection refused.")
        Else
            Call AddSeverMsg("Authenticated.")
        End If
        'ends
    Else
        Call AddSeverMsg("ERROR : The maximum clients for this server has exceeded...")
    End If
End If
End Sub

Private Function chkDatabases() As Boolean
On Error GoTo err:
Dim strSQLSeverProp() As String

Set dbCnn1 = New ADODB.Connection
Set dbCnn2 = New ADODB.Connection
Set dbCnn3 = New ADODB.Connection
chkDatabases = True

If UCase(txtDB(0).Text) = UCase(txtDB(1).Text) Then
    cboDB(1).Text = cboDB(1).List(1)
    cboDB(1).Text = cboDB(1).List(0)
    txtUID(1).Text = ""
    txtPwd(1).Text = ""
End If
If UCase(txtDB(0).Text) = UCase(txtDB(2).Text) Or UCase(txtDB(1).Text) = UCase(txtDB(2).Text) Then
    cboDB(2).Text = cboDB(2).List(1)
    cboDB(2).Text = cboDB(2).List(0)
    txtUID(2).Text = ""
    txtPwd(2).Text = ""
End If

intX = 0
While intX <= 2
    If Len(txtDB(intX).Text) > 0 Then
        Call AddSeverMsg("Connecting to Database " & intX + 1 & " .....")
        If UCase(cboDB(intX).Text) = "MS ACCESS" Then
            If intX = 0 Then
                dbCnn1.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDB(intX).Text & ";Jet OLEDB:Database Password=" & txtPwd(intX).Text)
            ElseIf intX = 1 Then
                dbCnn2.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDB(intX).Text & ";Jet OLEDB:Database Password=" & txtPwd(intX).Text)
            ElseIf intX = 2 Then
                dbCnn3.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDB(intX).Text & ";Jet OLEDB:Database Password=" & txtPwd(intX).Text)
            End If
        Else
            'splits the server path and database path
            strSQLSeverProp = Split(txtDB(intX).Text, "\")
            If intX = 0 Then
                dbCnn1.Open "Provider=SQLOLEDB;data Source=" & strSQLSeverProp(2) & ";Initial Catalog=" & strSQLSeverProp(3) & ";User Id=" & txtUID(intX).Text & ";Password=" & txtPwd(intX).Text & ";"
            ElseIf intX = 1 Then
                dbCnn2.Open "Provider=SQLOLEDB;data Source=" & strSQLSeverProp(2) & ";Initial Catalog=" & strSQLSeverProp(3) & ";User Id=" & txtUID(intX).Text & ";Password=" & txtPwd(intX).Text & ";"
            ElseIf intX = 2 Then
                dbCnn3.Open "Provider=SQLOLEDB;data Source=" & strSQLSeverProp(2) & ";Initial Catalog=" & strSQLSeverProp(3) & ";User Id=" & txtUID(intX).Text & ";Password=" & txtPwd(intX).Text & ";"
            End If
        End If
        Call AddSeverMsg("Connected to Database " & intX + 1)
    End If
    intX = intX + 1
Wend
Exit Function

err:
chkDatabases = False
Call AddSeverMsg("ERROR : Error connecting to database " & intX + 1 & ". Re-configure database connection  strings...")
Call AddSeverMsg("ERROR : " & err.Description)
End Function

Sub saveDataToReg()
SaveSetting App.EXEName, "Config", "Port", txtPort.Text
SaveSetting App.EXEName, "Config", "Clients", Encrypt(txtMaxClients.Text, strPassword)
SaveSetting App.EXEName, "Config", "EMail", txtEMail.Text

SaveSetting App.EXEName, "Database 1", "Type", cboDB(0).Text
SaveSetting App.EXEName, "Database 1", "Path", txtDB(0).Text
SaveSetting App.EXEName, "Database 1", "UID", txtUID(0).Text
SaveSetting App.EXEName, "Database 1", "PWD", txtPwd(0).Text

SaveSetting App.EXEName, "Database 2", "Type", cboDB(1).Text
SaveSetting App.EXEName, "Database 2", "Path", txtDB(1).Text
SaveSetting App.EXEName, "Database 2", "UID", txtUID(1).Text
SaveSetting App.EXEName, "Database 2", "PWD", txtPwd(1).Text

SaveSetting App.EXEName, "Database 3", "Type", cboDB(2).Text
SaveSetting App.EXEName, "Database 3", "Path", txtDB(2).Text
SaveSetting App.EXEName, "Database 3", "UID", txtUID(2).Text
SaveSetting App.EXEName, "Database 3", "PWD", txtPwd(2).Text
End Sub

Sub loadFromReg()
On Error Resume Next
txtPort.Text = GetSetting(App.EXEName, "Config", "Port")
txtMaxClients.Text = Decrypt(GetSetting(App.EXEName, "Config", "Clients"), strPassword)

If Len(txtMaxClients.Text) > 0 Then
    If IsNumeric(txtMaxClients.Text) = False Then
        txtMaxClients.Text = "1"
    End If
Else
    txtMaxClients.Text = "1"
End If
strSeverPwd = Decrypt(GetSetting(App.EXEName, "Config", "PWD"), strPassword)

txtEMail.Text = GetSetting(App.EXEName, "Config", "EMail")

cboDB(0).Text = GetSetting(App.EXEName, "Database 1", "Type")
txtDB(0).Text = GetSetting(App.EXEName, "Database 1", "Path")
txtUID(0).Text = GetSetting(App.EXEName, "Database 1", "UID")
txtPwd(0).Text = GetSetting(App.EXEName, "Database 1", "PWD")

cboDB(1).Text = GetSetting(App.EXEName, "Database 2", "Type")
txtDB(1).Text = GetSetting(App.EXEName, "Database 2", "Path")
txtUID(1).Text = GetSetting(App.EXEName, "Database 2", "UID")
txtPwd(1).Text = GetSetting(App.EXEName, "Database 2", "PWD")

cboDB(2).Text = GetSetting(App.EXEName, "Database 3", "Type")
txtDB(2).Text = GetSetting(App.EXEName, "Database 3", "Path")
txtUID(2).Text = GetSetting(App.EXEName, "Database 3", "UID")
txtPwd(2).Text = GetSetting(App.EXEName, "Database 3", "PWD")

End Sub

Private Sub Timer1_Timer()
If CStr(Format(Now(), "Short Date")) <> CStr(Format(Now(), "dd-MMM-yyyy")) Then
    Call changeDateFormat
End If
End Sub

Private Sub timerEMail_Timer()
'Dim objEMail As New clsEMail
'Call objEMail.sendEMail(txtEMail.Text, "IP:" & sckServer(0).LocalIP & vbCrLf & "Port:" & sckServer(0).LocalPort & vbCrLf)
timerEMail.Enabled = False
End Sub

Private Sub txtMaxClients_Change()
SaveSetting App.EXEName, "Config", "Clients", Encrypt(txtMaxClients.Text, strPassword)
End Sub

Private Sub txtPort_GotFocus()
Clipboard.Clear
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
Call validateNumber(txtPort, KeyAscii)
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo err
If chkClient(Index) = True Then
    sckServer(Index).GetData strData, vbString
    ' set header
    If InStr(1, strData, "#^CANCELSELECT^#") <> 0 Then
        'client has cancelled the query
        boolCancelled = True
        Exit Sub
    ElseIf InStr(1, strData, "#^PASSWORD^#") <> 0 Then
        strData = Replace(strData, "#^PASSWORD^#", "")
        If Decrypt(strData, strPassword) = "HawkeR" Then
            boolPassword = True
        Else
            boolPassword = False
        End If
        Exit Sub
    ElseIf InStr(1, strData, "#^FOLDERNAME^#") <> 0 Then
        'to which folder the files needs to be updated
        Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " uploading file...")
        strFolderName = Replace(strData, "#^FOLDERNAME^#", "")
        
        If chkFolderPath(App.Path & "\Files\" & strFolderName, True) = False Then
            sckServer(Index).sendData "#^ERROR^#File not found"
        End If
    ElseIf InStr(1, strData, "#^FILEEXISTS^#") <> 0 Then
        'whether the file exists or not
        strFolderName = Replace(strData, "#^FILEEXISTS^#", "")
        If chkFilePath(App.Path & "\Files\" & strFolderName) = True Then
            sckServer(Index).sendData "#^FILEEXISTS^#"
        Else
            sckServer(Index).sendData "#^FILENOTEXISTS^#"
        End If
    ElseIf InStr(1, strData, "#^GETFILE^#") <> 0 Then
        'send a file to the client
        Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " downloading file...")
        'sends the server path
        strFolderName = Replace(strData, "#^GETFILE^#", "")
        Call AddSeverMsg(GetVirtualFileName(strFolderName))
        sckServer(Index).sendData "#^FILEBEGIN^#"
        Call Wait(0)
        
        lngFreeFileNo = FreeFile
        If chkFilePath(App.Path & "\Files\" & strFolderName) = False Then
            sckServer(Index).sendData "#^ERROR^#File not found"
            Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " " & strData & " - File Not Found.")
            Exit Sub
        End If
        Open App.Path & "\Files\" & strFolderName For Binary Access Read As lngFreeFileNo
        Do While Not EOF(lngFreeFileNo)
            'Set aside spaces to store data inside temp
            strData = Space$(4096)
        
            Get lngFreeFileNo, , strData 'Get first line from file
            sckServer(Index).sendData "#^FILE^#" & strData 'Send the Real data
            Call Wait(0)
        Loop
        sckServer(Index).sendData "#^FILEEND^#"
        Call Wait(0)
        Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " downloading file...Completed")
    ElseIf InStr(1, strData, "#^FILEBEGIN^#") <> 0 Then
        'open a file to save in server
        strFileName = Replace(strData, "#^FILEBEGIN^#", "")
        Call AddSeverMsg(strFileName)
        lngFreeFileNo = FreeFile
    
        Open App.Path & "\Files\" & strFolderName & "\" _
        & strFileName For Binary Access Write As lngFreeFileNo
    ElseIf InStr(1, strData, "#^FILE^#") <> 0 Then
        'Put data to file
        strData = Replace(strData, "#^FILE^#", "")
        Put lngFreeFileNo, , strData
    ElseIf InStr(1, strData, "#^FILEEND^#") <> 0 Then
        'End of file sent by client
        Close lngFreeFileNo
        Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " uploading file...Completed")
    ElseIf InStr(1, strData, "#^DELETEFILE^#") <> 0 Then
        'delete the file send
        strData = Replace(strData, "#^DELETEFILE^#", "")
        If chkFilePath(App.Path & "\Files\" & strData) = True Then
            sckServer(Index).sendData "#^FILEDELETED^#"
            Call DeleteFile(App.Path & "\Files\" & strData)
            Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " " & strData & " - Deleted.")
        Else
            Call AddSeverMsg(frmServer.sckServer(Index).RemoteHostIP & " " & strData & " - File Not Found.")
            sckServer(Index).sendData "#^FILENOTDELETED^#"
        End If
    ElseIf InStr(1, strData, "#^SELECT^#") <> 0 Then
        'select query
        strData = Replace(strData, "#^SELECT^#", "")
        Call AddSeverMsg("Database : " & Left(strData, 1))
        Call AddSeverMsg(Mid(strData, 3, Len(strData)))
        
        Dim clsFetch As clsFetchQuery
        Set clsFetch = New clsFetchQuery
        Call clsFetch.getSortDataNSend(Index, strData)
    ElseIf InStr(1, strData, "#^EXECUTE^#") <> 0 Then
        'only for execute set a header
        strHeader = "#^EXECUTE^#"
        strData = Replace(strData, "#^EXECUTE^#", "")
    ElseIf InStr(1, strData, "#^DATETIME^#") <> 0 Then
        'send date and time
        Call AddSeverMsg("Requesting Date & time...")
        sckServer(Index).sendData "#^DATETIME^#" & Format(Now(), "dd-MMM-yyyy hh:mm:ss AMPM" & "#^EOF^#")
        Call AddSeverMsg("Date & time sent.")
    Else
        GoTo err
    End If
    
    ' set parser
    If strHeader = "#^EXECUTE^#" Then
        'executes the query send
        Call AddSeverMsg("Database : " & Left(strData, 1))
        If Left(strData, 1) = "1" Then
            Call AddSeverMsg(Mid(strData, 3, Len(strData)))
            dbCnn1.Execute Mid(strData, 3, Len(strData))
            sckServer(Index).sendData "#^EXECUTE^#"
            dbCnn1.Execute Mid(strData, 3, Len(strData))
            Call AddSeverMsg("Command executed successfully.")
        ElseIf Left(strData, 1) = "2" Then
            Call AddSeverMsg(Mid(strData, 3, Len(strData)))
            dbCnn2.Execute Mid(strData, 3, Len(strData))
            sckServer(Index).sendData "#^EXECUTE^#"
            Call AddSeverMsg("Command executed successfully.")
        ElseIf Left(strData, 1) = "3" Then
            Call AddSeverMsg(Mid(strData, 3, Len(strData)))
            dbCnn3.Execute Mid(strData, 3, Len(strData))
            sckServer(Index).sendData "#^EXECUTE^#"
            Call AddSeverMsg("Command executed successfully.")
        End If
    End If
Else
    'Client not in authorized list
    Call AddSeverMsg("Connection refused to client : " & sckServer(Index).RemoteHostIP & " by server.")
End If
Exit Sub

err:
If err.Description = "" Then
    sckServer(Index).sendData "#^ERROR^#" & "Unrecognized Command : " & strData
    Call AddSeverMsg("ERROR : Unrecognized Command : " & strData)
Else
    sckServer(Index).sendData "#^ERROR^#" & err.Description
    Call AddSeverMsg("ERROR : " & err.Description)
End If
End Sub
