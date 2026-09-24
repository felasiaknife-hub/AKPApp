VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VideoSoft VSFlexGrid Pro 7.0"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   810
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Frozen Cells"
      Height          =   495
      Index           =   16
      Left            =   10140
      TabIndex        =   16
      Tag             =   "Implements frozen columns and rows."
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Explorer"
      Height          =   495
      Index           =   15
      Left            =   10140
      TabIndex        =   15
      Tag             =   "Implements the VSFlexGrid as a File Explorer."
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Owner&Draw"
      Height          =   495
      Index           =   5
      Left            =   10140
      TabIndex        =   13
      Tag             =   "Invoke Windows APIs to customize the look of any cell."
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "C&alendar"
      Height          =   495
      Index           =   7
      Left            =   10140
      TabIndex        =   10
      Tag             =   "Use the VSFlexGrid as a calendar."
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Balance Sheet"
      Height          =   495
      Index           =   14
      Left            =   10140
      TabIndex        =   7
      Tag             =   "Customize the VSFlexGrid control to behave like a powerful outliner."
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "HTML"
      Height          =   495
      Index           =   11
      Left            =   9000
      TabIndex        =   14
      Tag             =   "Save the content of the VSFlexGrid into HTML."
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Lis&ts"
      Height          =   495
      Index           =   12
      Left            =   9000
      TabIndex        =   12
      Tag             =   "Dress your applications with custom listboxes."
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&FlexGroup"
      Height          =   495
      Index           =   13
      Left            =   9000
      TabIndex        =   9
      Tag             =   "Add MS Outlook-style grouping and sorting."
      Top             =   4200
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&SubTotals"
      Height          =   495
      Index           =   4
      Left            =   9000
      TabIndex        =   6
      Tag             =   "Automatically calculate subtotals across multiple columns."
      Top             =   3720
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&XML"
      Height          =   495
      Index           =   3
      Left            =   7860
      TabIndex        =   11
      Tag             =   "Parse XML files and load them into the grid."
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Invoice"
      Height          =   495
      Index           =   6
      Left            =   7860
      TabIndex        =   8
      Tag             =   "Combine multiple VSFlexGrids for complex data entry forms."
      Top             =   4200
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&MergeCells"
      Height          =   495
      Index           =   1
      Left            =   7860
      TabIndex        =   5
      Tag             =   "Merge cells and pivot data within your application."
      Top             =   3720
      Width           =   1155
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&About"
      Height          =   495
      Index           =   10
      Left            =   5760
      TabIndex        =   4
      Tag             =   "Contacting VideoSoft."
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Features"
      Height          =   495
      Index           =   8
      Left            =   5760
      TabIndex        =   3
      Tag             =   "Learn how to implement the key features of VSFlexGrid Pro 7.0"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "What's New"
      Height          =   495
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Tag             =   "What's new in VSFlexGrid Pro 7.0?"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Comparison"
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   2
      Tag             =   "Feature Comparison Chart."
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Overview"
      Height          =   495
      Index           =   9
      Left            =   4560
      TabIndex        =   0
      Tag             =   "Product description and highlights."
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Height          =   675
      Left            =   4320
      TabIndex        =   17
      Top             =   5880
      Width           =   5715
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   240
      Picture         =   "frmMenu.frx":0000
      Top             =   60
      Width           =   8550
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   3780
      Picture         =   "frmMenu.frx":79F0
      Top             =   1860
      Width           =   7740
   End
   Begin VB.Image Image4 
      Height          =   375
      Index           =   1
      Left            =   8040
      Picture         =   "frmMenu.frx":CA31
      Top             =   3000
      Width           =   705
   End
   Begin VB.Image Image4 
      Height          =   375
      Index           =   0
      Left            =   4560
      Picture         =   "frmMenu.frx":D37E
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   7800
      X2              =   11040
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4440
      X2              =   7200
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image imgLogo 
      Height          =   3840
      Left            =   60
      Picture         =   "frmMenu.frx":E32D
      Top             =   2940
      Width           =   4200
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
    
    MousePointer = 11
    
    Select Case Index
    
    
        Case 0: frm7High.Show
        Case 1: frmMerge.Show
        Case 2: frmComp.Show
        Case 3: frmXML.Show
        Case 4: frmTotals.Show
        Case 5: frmOwner.Show
        Case 6: frmInvoic.Show
        Case 7: frmCal.Show
        Case 8: frmNew6.Show
        Case 9: frmOverall.Show
        Case 10
                MsgBox "VideoSoft                                                              " & vbCrLf & _
                       "5900-T Hollis Street" & vbCrLf & _
                       "Emeryville, CA 94608" & vbCrLf & vbCrLf & _
                       "Voice: 510-595-2400" & vbCrLf & _
                       "Fax:   510-595-2424" & vbCrLf & vbCrLf & _
                       "www.videosoft.com", vbOKOnly, "Contacting VideoSoft"
                       
        Case 11: frmHTML.Show
        Case 12: frmList.Show
        Case 13: frmGroup.Show
        Case 14: fSheet.Show
        Case 15: frmExplo.Show
        Case 16: frmfreez.Show
        
        
        
        Case Else
            MsgBox "Sorry, this demo is not available.", vbOKOnly, App.Title
        
    End Select
    
    MousePointer = 0
    
End Sub


Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i%
    
    On Error Resume Next
    
    For i = 0 To 17
        cmd(i).FontBold = False
    Next
    
    cmd(Index).FontBold = True
    lblHelp = cmd(Index).Tag
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i%
    On Error Resume Next
    
    For i = 0 To 14
        cmd(i).FontBold = False
    Next
    
    lblHelp = ""
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    End
    
End Sub


