VERSION 5.00
Begin VB.Form frmCancelDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel Detail"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   6960
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   5160
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtCancelDetail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   8415
   End
   Begin VB.Label Label1 
      Caption         =   "รายละเอียดการยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmCancelDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTxt As String, strTmp As String

Private Sub cmdCancel_Click()
    strTxt = ""
    ValidateText True
End Sub

Private Sub cmdOK_Click()
    strTxt = txtCancelDetail.Text
    ValidateText True
End Sub

Public Function CancelText(Optional ByVal strCancel As String = "") As String
    strTmp = strCancel
    Call ValidateText
    CancelText = strTxt
End Function

Private Sub Form_Activate()
    txtCancelDetail.Text = strTmp
End Sub

Private Sub ValidateText(Optional ByVal Activate As Boolean = False)
On Error Resume Next
    If Activate = False Then
        Me.ZOrder
        Me.Show 1
    Else
        Unload Me
    End If
End Sub
