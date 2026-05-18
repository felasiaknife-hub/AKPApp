VERSION 5.00
Begin VB.Form frmFormDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "เลือกฟอร์มเอกสารที่ต้องการพิมพ์"
   ClientHeight    =   2880
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   465
      Left            =   4095
      TabIndex        =   14
      Top             =   1215
      Width           =   1815
      Begin VB.CheckBox Check2 
         Caption         =   "ลายน้ำ"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "เลขทะเบียน Lab"
      Height          =   750
      Left            =   4095
      TabIndex        =   9
      Top             =   60
      Width           =   1845
      Begin VB.OptionButton OptRegno 
         Caption         =   "มีเลขทะเบียน"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptNoReg 
         Caption         =   "ไม่มีเลขทะเบียน"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "เลือกฟอร์ม"
      Height          =   1740
      Left            =   75
      TabIndex        =   4
      Top             =   60
      Width           =   3915
      Begin VB.OptionButton Option7 
         Caption         =   "ผล X-Ray"
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Extraction Test"
         Height          =   255
         Left            =   270
         TabIndex        =   8
         Top             =   900
         Width           =   3135
      End
      Begin VB.OptionButton Option5 
         Caption         =   "พิมพ์ผลการวิเคราะห์องค์ประกอบของเสีย"
         Height          =   255
         Left            =   270
         TabIndex        =   7
         Top             =   585
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "พิมพ์ผลการวิเคราะห์องค์ประกอบของเสีย BME"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   3645
      End
      Begin VB.OptionButton Option2 
         Caption         =   "พิมพ์ Waste Data ฟอร์ม อัคคีปราการ"
         Height          =   255
         Left            =   2700
         TabIndex        =   5
         Top             =   135
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "เลือกประเภทการพิมพ์"
      Height          =   975
      Left            =   75
      TabIndex        =   1
      Top             =   1800
      Width           =   3915
      Begin VB.OptionButton Option4 
         Caption         =   "พิมพ์ออกเครื่องพิมพ์"
         Height          =   195
         Left            =   270
         TabIndex        =   3
         Top             =   660
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "พิมพ์ออกหน้าจอ"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   600
      Left            =   4485
      TabIndex        =   0
      Top             =   1755
      Width           =   1395
   End
   Begin VB.Frame Frame4 
      Height          =   465
      Left            =   4095
      TabIndex        =   12
      Top             =   765
      Width           =   1815
      Begin VB.CheckBox Check1 
         Caption         =   "ใช้ฟอร์มเก่า"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmFormDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormType As Integer

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Dim i As Integer
    Dim UserGrp As New AWSUserGroup
            If Not CurrentUser Is Nothing Then
            For i = 1 To CurrentUser.getUserGroups.Count
                Set UserGrp = CurrentUser.getUserGroups(i)
                If Trim(UserGrp.Name) = "Environment" Or Trim(UserGrp.Name) = "Lab" Or Trim(UserGrp.Name) = "Senior LAB" Then
                    Option5.Enabled = True
                    Option6.Enabled = True
                    Else
                    Option5.Enabled = False
                    Option6.Enabled = False
                End If
            Next i
        End If
End Sub
Private Sub OKButton_Click()
Dim PrintType As Integer
    If Option1.Value = True Then
        FormType = 0
        frmWasteDataQue.RptAnaLysisNo = 3
    If OptRegno.Value = True Then frmWasteDataQue.RptAnaLysisNo = 4
    If OptNoReg.Value = True Then frmWasteDataQue.RptAnaLysisNo = 5
    If OptRegno.Value = True And Check1.Value = 1 Then frmWasteDataQue.RptAnaLysisNo = 6
    If OptNoReg.Value = True And Check1.Value = 1 Then frmWasteDataQue.RptAnaLysisNo = 7
    ElseIf Option2.Value = True Then
        FormType = 1
        frmWasteDataQue.RptAnaLysisNo = 0
    ElseIf Option5.Value = True Then
        FormType = 1
        frmWasteDataQue.RptAnaLysisNo = 1
    If OptRegno.Value = True Then frmWasteDataQue.RptAnaLysisNo = 4
    If OptNoReg.Value = True Then frmWasteDataQue.RptAnaLysisNo = 5
    If OptRegno.Value = True And Check1.Value = 1 Then frmWasteDataQue.RptAnaLysisNo = 6
    If OptNoReg.Value = True And Check1.Value = 1 Then frmWasteDataQue.RptAnaLysisNo = 7
    ElseIf Option6.Value = True Then
        FormType = 1
        frmWasteDataQue.RptAnaLysisNo = 2
    ElseIf Option7.Value = True Then
        FormType = 1
        frmWasteDataQue.RptAnaLysisNo = 8
    Else
        FormType = 0
    End If
    If Option3.Value = True Then PrintType = 0
    If Option4.Value = True Then PrintType = 1
    
    If Check2.Value = 1 Then frmWasteDataQue.Watermark = True
    If Check2.Value = 0 Then frmWasteDataQue.Watermark = False
    
    frmWasteDataQue.FormType = FormType
    frmWasteDataQue.PrintType = PrintType
    Unload Me
End Sub

