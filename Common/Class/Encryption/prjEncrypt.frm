VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   525
      Left            =   2310
      TabIndex        =   7
      Top             =   1650
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   525
      Left            =   990
      TabIndex        =   6
      Top             =   1650
      Width           =   1245
   End
   Begin VB.TextBox txtResult 
      Height          =   315
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1020
      Width           =   2805
   End
   Begin VB.TextBox txtSrc 
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Top             =   570
      Width           =   2805
   End
   Begin VB.TextBox txtKey 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   150
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Result"
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   4
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Src Text"
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Key"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim tmpEnc As RC4Crypt
    If Trim(txtKey.Text) = "" Then
        MsgBox "Please enter key"
        Exit Sub
    End If
    If Trim(txtSrc.Text) = "" Then
        MsgBox "Please enter src text"
        Exit Sub
    End If
    Set tmpEnc = New RC4Crypt
    txtResult.Text = tmpEnc.EncryptString(txtSrc.Text, txtKey.Text)
    Set tmpEnc = Nothing
End Sub

Private Sub Command2_Click()
Dim tmpEnc As RC4Crypt
    If Trim(txtKey.Text) = "" Then
        MsgBox "Please enter key"
        Exit Sub
    End If
    If Trim(txtSrc.Text) = "" Then
        MsgBox "Please enter src text"
        Exit Sub
    End If
    Set tmpEnc = New RC4Crypt
    txtResult.Text = tmpEnc.DecryptString(txtSrc.Text, txtKey.Text)
    Set tmpEnc = Nothing
End Sub
