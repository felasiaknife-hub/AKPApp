VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Delete File"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get File"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "select * from 1"
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Put File"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Execute Query"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Date Time"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Execute:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRec As New ADODB.Recordset

Private Sub Command1_Click()
Call checkDateTime
If boolError = False Then
    Text1.Text = Text1.Text & vbCrLf & strNowDateTime
End If
End Sub

Private Sub Command2_Click()
Call fetchData(Text2.Text, 1)
If boolError = False Then
    Set rsRec = rsSelect.Clone
    Text1.Text = Text1.Text & vbCrLf & rsRec.RecordCount
    Text1.Text = Text1.Text & vbCrLf & rsRec(0).Value
Else
    Text1.Text = Text1.Text & vbCrLf & "Error" & vbCrLf & strServerError
End If
End Sub

Private Sub Command3_Click()
Call executeQuery(Text2.Text, 1)
If boolError = False And boolQueryExecuted = True Then
    Text1.Text = Text1.Text & vbCrLf & "Command executed successfully."
Else
    Text1.Text = Text1.Text & vbCrLf & "Error in command"
End If
End Sub

Private Sub Command4_Click()
Call sendFile(App.Path & "\Winter.jpg", "MY")
End Sub

Private Sub Command5_Click()
Call ReceiveFile("My\Winter.jpg", App.Path)
End Sub

Private Sub Command6_Click()
Call DeleteFile("\My\Winter.jpg")
End Sub

Private Sub Form_Load()
Load frmClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmClient
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub
