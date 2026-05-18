VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWasteDataQue 
   BackColor       =   &H00C0C0FF&
   Caption         =   "กากที่รอการตรวจรับ"
   ClientHeight    =   10050
   ClientLeft      =   1320
   ClientTop       =   30
   ClientWidth     =   11400
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CRrpt 
      Left            =   2880
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4260
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":0296
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":02E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":05B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":087F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":0B4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   4260
      Top             =   4230
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   10050
      Left            =   0
      ScaleHeight     =   10020
      ScaleWidth      =   2355
      TabIndex        =   15
      Top             =   0
      Width           =   2385
      Begin VB.CommandButton cmdPrintAnalysis 
         Caption         =   "พิมพ์ข้อมูลวิเคราะห์"
         Height          =   400
         Left            =   60
         TabIndex        =   29
         Top             =   7080
         Width           =   2205
      End
      Begin VB.ComboBox cboWaste 
         Height          =   315
         ItemData        =   "frWasteDataQue.frx":0DCB
         Left            =   180
         List            =   "frWasteDataQue.frx":0DCD
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintSum 
         Caption         =   "พิมพ์ข้อมูลสรุป"
         Height          =   400
         Left            =   60
         TabIndex        =   25
         Top             =   6600
         Width           =   2205
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         ItemData        =   "frWasteDataQue.frx":0DCF
         Left            =   180
         List            =   "frWasteDataQue.frx":0DD1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ไม่เลือกเงื่อนไข"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   3360
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "พิมพ์ข้อมูลเวสต์"
         Height          =   400
         Left            =   60
         TabIndex        =   12
         Top             =   6120
         Width           =   2205
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เฉพาะรายการที่ยังไม่มีวันหมดอายุใบอนุญาต"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   4200
         Width           =   1995
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เฉพาะรายการที่ยังไม่ออกเลข กอ."
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   3720
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เฉพาะกากฯ ที่ใบอนุญาตจะหมดอายุ ในอีก 30 วัน"
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   180
         TabIndex        =   10
         Top             =   5160
         Width           =   2115
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เฉพาะกากฯ ที่จะหมดอายุในอีก ไม่เกิน 30 วัน"
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   4680
         Width           =   2055
      End
      Begin WasteDataForm.ctlDate dtTo 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   2880
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteDataForm.ctlDate dtFrom 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   2160
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   400
         Left            =   60
         TabIndex        =   13
         Top             =   7560
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ค้นหา"
         Height          =   400
         Left            =   60
         TabIndex        =   11
         Top             =   5640
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frWasteDataQue.frx":0DD3
         Left            =   180
         List            =   "frWasteDataQue.frx":0DD5
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   2055
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lbOpen1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "รอตรวจรับ"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   8040
         Width           =   1095
      End
      Begin VB.Label lbWaste1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ต้องการ ปรับ Waste"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   27
         Top             =   9000
         Width           =   1395
      End
      Begin VB.Shape Shape1_1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   8
         Left            =   120
         Top             =   9000
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   120
         Top             =   9720
         Width           =   195
      End
      Begin VB.Label lblQApprove 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รออนุมัติใบเสนอราคา"
         ForeColor       =   &H00C0C000&
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   9720
         Width           =   1515
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   5
         Left            =   120
         Top             =   8280
         Width           =   195
      End
      Begin VB.Label lblApproved1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจรับแล้ว, ยังไม่สรุปผล"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   8280
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   120
         Top             =   9480
         Width           =   195
      End
      Begin VB.Label lblReject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตีกลับ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   9480
         Width           =   405
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   9240
         Width           =   495
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจสอบแล้ว, ไม่รับกำจัด"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   8760
         Width           =   1815
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจสอบแล้ว, รับกำจัด"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   8520
         Width           =   1635
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   120
         Top             =   9240
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   120
         Top             =   8760
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   120
         Top             =   8520
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   120
         Top             =   8040
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   2580
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตั้งแต่วันที่"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อของเสีย, ชื่อลูกค้า :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   1515
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2280
      TabIndex        =   14
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "วันที่ส่งข้อมูล"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ตรวจรับ"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อกากฯ"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เลข 6 หลัก"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "สถานะ"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ใบอนุญาต"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "เลข กอ."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "วิธีกำจัด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ผล LAB"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "isNew"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Sales"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "หมายเหตุ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "บริษัทที่รับกำจัด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ปี"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "วันที่รับตัวอย่าง"
         Object.Width           =   2787
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "วันที่วิเคราะห์"
         Object.Width           =   2787
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "ผลวิเคราะห์"
         Object.Width           =   2892
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "การอนุมัติผลวิเคราะห์"
         Object.Width           =   2892
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "ใหม่/เก่า"
         Object.Width           =   2540
      EndProperty
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรุ่น,ลาย หรือยี่ห้อ :"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   17
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmWasteDataQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colDealerGroup_ As Collection
Dim colCurShowRoom_ As Collection
Dim curCondition_ As String
Dim isEnv As Boolean, isIT As Boolean, isLab As Boolean
Public Watermark As Boolean
Public FormType As Integer '*** 1 is AP else is BWG, BWC, ET
Public PrintType As Integer  '0 is defaul, 1 is printer
Public RptAnaLysisNo As Integer
Dim Sql1 As String, Sql2 As String   'For print report waste request
Private Sub SetUserGroup()
Dim cVal As Boolean
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim i As Integer
    isEnv = False: isIT = False: isLab = False
    Set UserGrp = CurrentUser.getUserGroups
    For i = 1 To UserGrp.Count
        Set Group = UserGrp(i)
        If Group.Name = "Administrator" Then
            StrGrp = Group.Name
            isIT = True: isEnv = False
            Exit For
        ElseIf LCase(Group.Name) = "environment" Then
            StrGrp = Group.Name
            isEnv = True
        ElseIf LCase(Group.Name) = "lab" Then
            StrGrp = Group.Name
            isLab = True
        End If
    Next
End Sub
Private Sub PrintRpt(ByVal wastedataID As String, ByVal PrintType As Integer)
    With CRrpt
        .Reset
        If FormType = 1 Then
            .ReportFileName = App.Path & "\Reports\WasteDataFormAP.rpt"
        Else
            .ReportFileName = App.Path & "\Reports\WasteDataForm.rpt"
        End If
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .WindowTitle = "ข้อมูลกากอุตสาหกรรม"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .SelectionFormula = "{vw_WasteDataCRSearch.WasteDataID} = '" & Replace(wastedataID, "'", "") & "'"
        .Destination = PrintType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
Private Sub PrintRptAnalysis(ByVal wastedataID As String, ByVal PrintType As Integer)
   With CRrpt
        .Reset
        If RptAnaLysisNo = 0 Then
        .ReportFileName = App.Path & "\Reports\WastedataAnalysis.rpt"
        ElseIf RptAnaLysisNo = 1 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\WasteAnalysisap.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\WasteAnalysisapa.rpt"
            End If
        ElseIf RptAnaLysisNo = 2 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\ExtractionReport.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\ExtractionReportA.rpt"
            End If
        ElseIf RptAnaLysisNo = 4 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNewA.rpt"
            End If
        ElseIf RptAnaLysisNo = 5 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew2.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew2A.rpt"
            End If
        ElseIf RptAnaLysisNo = 6 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew3.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew3A.rpt"
            End If
        ElseIf RptAnaLysisNo = 7 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew4.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew4A.rpt"
            End If
        ElseIf RptAnaLysisNo = 8 Then
            If Watermark = True Then
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew5A.rpt"
            Else
                .ReportFileName = App.Path & "\Reports\WasteAnalysisAPNew5.rpt"
            End If
        Else
        .ReportFileName = App.Path & "\Reports\WasteAnalysis.rpt"
        End If
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .WindowTitle = "ข้อมูลกากอุตสาหกรรม"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .SelectionFormula = "{vw_WasteDataCRSearch.WasteDataID} = '" & Replace(wastedataID, "'", "") & "'"
        .Destination = PrintType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim wastedataID As String, i As Integer
    Load frmFormDialog
    frmFormDialog.Show 1
    For i = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(i).Checked = True Then
            wastedataID = lvResult.ListItems(i).Key
            PrintRpt wastedataID, PrintType
        End If
    Next
End Sub

Private Sub cmdPrintAnalysis_Click()
Dim wastedataID As String, i As Integer
Dim tSql As String
Dim RS As ADODB.Recordset

    For i = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(i).Checked = True Then
            wastedataID = lvResult.ListItems(i).Key
            tSql = "SELECT isApproveLab FROM tbWasteDataLab where WasteDataID=" & wastedataID & " "
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
                If RS!isApproveLab = "Y" Or isLab = True Then
                    frmFormDialog.Check2.Value = 1
                    Load frmFormDialog
                    frmFormDialog.Show 1
                    PrintRptAnalysis wastedataID, PrintType
                Else
                    MsgBox "คุณไม่สามารถพิมพ์ข้อมูลวิเคราะห์ที่ยังไม่อนุมัติผลได้ กรุณาติดต่อ LAB", vbInformation
                End If
        End If
    Next
End Sub

Private Sub cmdPrintSum_Click()
Dim tSql As String
    cmdSearch_Click
    tSql = Sql1
    DBManager.GetDBConnection.Execute "Alter View vw_WasteDataCRSearch_Print  AS " & tSql
    tSql = "SELECT * FROM vw_WasteDataCRSearch_Print"
    With CRrpt
        .Reset
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\WasteRequest.rpt"
        .WindowTitle = "ข้อมูลเวสต์ของลูกค้าที่รอการตรวจรับ"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
'        .SelectionFormula = "{vw_WasteDataCRSearch.WasteDataID} = '" & Replace(WasteDataID, "'", "") & "'"
        .SQLQuery = tSql
        .Destination = PrintType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdSearch_Click()
If dtFrom.YearValue <> dtTo.YearValue Then
    MsgBox "คุณสามารถค้นหา ข้อมูล กากอุตสาหกรรม ได้ในปี พ.ศ. เดียวกันเท่านั้น", vbInformation
    Exit Sub
End If
 SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.ActiveControl.Container.Name = txtKey.Container.Name Or Me.ActiveControl.Container.Name = cboJobStatus.Container.Name Then
            If Not TypeOf Me.ActiveControl Is CommandButton Then
                cmdSearch_Click
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim tmpRec As New ADODB.Recordset
Dim backDate As String
Dim tmpY As Integer, tmpM As Integer, tmpD As Integer
Dim tmpArr
    
'    tmpRec.Open "select backdate=convert(varchar(10),getdate()-7,111)", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
'    backDate = Trim("" & tmpRec("BackDate"))
'    tmpRec.Close
    SetUserGroup
    PopulateCompany
    If IsDate(backDate) Then
        tmpArr = Split(backDate, "/")
        If CInt(tmpArr(0)) < 2500 Then
            dtFrom.ValueYMD = Trim("" & (CInt(tmpArr(0) + 543))) & "/" & Trim("" & tmpArr(1)) & "/" & Trim("" & tmpArr(2))
        Else
            dtFrom.ValueYMD = backDate
        End If
    End If
    Set tmpRec = Nothing
    With cboJobStatus
        .Clear
        .AddItem "== ทุกสถานะ =="
        .AddItem "ยังไม่ส่งข้อมูลให้ Env"
        .AddItem "รอตรวจรับ"
        .AddItem "ยังไม่สรุปผล"
        .AddItem "ตีกลับ"
        .AddItem "รับกำจัด"
        .AddItem "ไม่รับกำจัด"
        .AddItem "ยกเลิก"
        .AddItem "รออนุมัติใบเสนอราคา"
        .AddItem "Waste ที่ต้องการ Rewrie"
        .AddItem "Waste ที่ถูก Revised"
        .AddItem "ส่งวิเคราะห์ (AKP)"
        .AddItem "รอผลวิเคราะห์ (AKP)"
        .AddItem "วิเคราห์แล้ว(AKP)"
        .ListIndex = 0
    End With
       
    With cboWaste
        .Clear
        .AddItem "== แสดงทั้งหมด =="
        .AddItem "แสดงเวสต์ที่ต้องตรวจสอบ"
        .AddItem "แสดงเวสต์ที่ไม่ต้องตรวจสอบ"
        .ListIndex = 1
    End With
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    cboCompany.AddItem "== เลือกบริษัท =="
    For x = 1 To CurrentUser.WasteCompany.Count
       If CurrentUser.WasteCompany(x).isDeleted <> "Y" Then
        cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyCode
        End If
        'If CurrentUser.WasteCompany(x).CompanyCode = "AKP" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyCode
        'If CurrentUser.WasteCompany(x).CompanyCode = "BWG" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyCode
        'If CurrentUser.WasteCompany(x).CompanyCode = "BWC" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyCode

    Next
    cboCompany.ListIndex = 0
End Sub
Private Sub Form_Resize()
On Error Resume Next
    With lvResult
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 100
        End If
        .Height = Me.Height - .Top - 410
    End With
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'MsgBox ColumnHeader.Width
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If lvResult.SelectedItem Is Nothing Then Exit Sub
'       If lvResult.SelectedItem.ListSubItems(10).Text = "Y" Then
'        Set tmpCol = tmpS.CustomerSearch("CustomerID in (Select CustomerID from tbWasteData_New Where WasteDataID=" & lvResult.SelectedItem.Key & ")")
'    Else
        Set tmpCol = tmpS.CustomerSearch("CustomerID in (Select CustomerID from tbWasteDataCR Where WasteDataID=" & lvResult.SelectedItem.Key & ")")
'    End If
    If tmpCol.Count > 0 Then
        Load frmWasteData
        With frmWasteData
            .SetCallerForm Me, tmpCol(1)
            .SetDataWasteData tmpCol(1).WasteItems(lvResult.SelectedItem.Key)
            .Show 1
        End With
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub Timer1_Timer()
    'SearchQue
End Sub

Private Sub SearchQue()
'On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, Y As Long, i As Integer
Dim sqlStr As String, sqlStr1 As String, sqlStr2 As String, sqlStr3 As String, strCond As String, statusDesc As String
Dim tmpArr
Dim Company As String
Dim curColor
Dim subStr As String, LastYear As String, CurYear As String
Dim TotalCount As Long, LastY As Integer, OneMonth As String, Reject As Boolean
Dim NewCount As Long
NewCount = 0
    LastY = Left(TodayDate, 4) - 1
    LastYear = LastY & Mid(TodayDate, 5)
    OneMonth = LastY & Format(Date + 30, "/mm/dd")
    CurYear = Format(Date + 30, "yyyy/mm/dd")
    sqlStr = "Select WasteDataID,SaleSendDateAuto,AcceptedDate,WasteNo,WasteName, '[' + Customercode  + ']' + CompanyName as CompanyName,EUCode1,EUCode2,EUCode3," & vbCrLf
    sqlStr = sqlStr & "isApproveByLab,WasteType,isTreatAble,isCancel,LicenseExpDate,KorOrNumber," & vbCrLf
    sqlStr = sqlStr & "TreatmentDesc,isLabResultNeeded, ISNULL(isRequest, 'N') AS isRequest, ISNULL(isAcceptedByEnvi, '') AS isAcceptedByEnvi, " & vbCrLf
    sqlStr = sqlStr & "ISNULL(isNew, 'N') AS isNew, SaleStaff, ISNULL(RejectDetail, 'N') AS RejectDetail, ISNULL(CompanyCode, '') AS CompanyCode, " & vbCrLf
    sqlStr = sqlStr & "RequestDate, Status, isQuotationReq, isApprove, CisNew, CisRequest, RejectDesc,WasteYears,CStatus,LBstatus,CreatedDate,ISnull(CaseName,'N') as CaseName, " & vbCrLf
    sqlStr = sqlStr & "ReceivedDate, AnalyticalDate, LBStatus as LabStatus, isApproveLab,WasteReferrent" & vbCrLf
    sqlStr1 = "From vw_WasteDataCRSearch " & vbCrLf
'    sqlStr2 = "From vw_WasteDataCR_NewSearch " & vbCrLf
    If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
        strCond = "WHERE  (WasteName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or Customercode Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%') And (isCancel <> 'Y') " & vbCrLf
    End If
    If Option3.Value = True Or Option4.Value = True Then
    Else
        If InStr(1, LCase(strCond), "where") > 0 Then
            strCond = strCond & " And " & subStr
        Else
            strCond = strCond & " Where " & subStr
        End If
        If cboJobStatus.ListIndex = 1 Then
            strCond = strCond & " (CreatedDate >= '" & dtFrom.ValueYMD & "' And CreatedDate <= '" & dtTo.ValueYMD & "') " & vbCrLf
        Else
            strCond = strCond & " (SaleSendDateAuto >= '" & dtFrom.ValueYMD & "' And SaleSendDateAuto <= '" & dtTo.ValueYMD & "') " & vbCrLf
        End If
    End If
    
    Select Case cboJobStatus.ListIndex
       Case 1 'ยังไม่ส่งข้อมูลให้ Env
            subStr = " (Status = 'D') "
        Case 2 'รอตรวจรับ
            subStr = " (Status = 'R') "
        Case 3 'ยังไม่สรุปผล
            subStr = " (Status = 'P') "
        Case 4 'ตีกลับ
            subStr = " (Status = 'J') "
        Case 5 'บำบัดได้
            subStr = " (Status = 'A') "
        Case 6 'บำบัดไม่ได้
            subStr = " (Status = 'E') "
        Case 7 'ยกเลิก
            subStr = " (Status='C') "
        Case 8 'รออนุมัติ'
            subStr = " (Status = 'W') "
        Case 9
            subStr = " (Status = 'AR') "
        Case 10
            subStr = " (Status = 'IA') "
        Case 11 'ส่งวิเคราะห์
            subStr = "(Status = 'P' And LBstatus = 'S')"
        Case 12 'รอวิเคราะห์
            subStr = "(Status = 'P' And LBstatus = 'R')"
        Case 13 'วิเคราะห์แล้ว
            subStr = "((Status ='P' or Status='E' or Status = 'A')  And LBstatus = 'A') "
    End Select
    
    If cboJobStatus.ListIndex > 0 Then
        If InStr(1, LCase(strCond), "where") > 0 Then
            strCond = strCond & " And " & subStr
        Else
            strCond = strCond & " Where " & subStr
        End If
    End If
    If cboWaste.ListIndex = 1 Then
'        If (cboJobStatus.ListIndex > 0) Then
'            If Trim(subStr) <> "(Status = 'R') OR (Status ='W')" Then
                strCond = strCond & " AND (ISNULL(isQuotationReq, 'N') <> 'Y') "
'            End If
'        End If
    ElseIf cboWaste.ListIndex = 2 Then
        strCond = strCond & " AND (ISNULL(isQuotationReq, 'N') = 'Y') "
    End If
    If cboCompany.ListIndex > 0 Then
        strCond = strCond & " AND (CompanyCode = '" & Trim(cboCompany.Text) & "')"
    End If
'    If isAkp = True And isBWG = False And isPST = False Then
'        strCond = strCond & " And (CompanyID = '776EDA5830F744A6B3551348D851BC22') "
'    ElseIf isAkp = False And isBWG = False And isPST = True Then
'        strCond = strCond & " And (SDName like 'PST%') "
'    End If
       
    If Option5.Value = False Then
        If Option1.Value = True Then
            If InStr(1, LCase(strCond), "where") > 0 Then
                strCond = strCond & " And KorOrNumber='' "
            Else
                strCond = strCond & " Where KorOrNumber='' "
            End If
        End If
        
        If Option2.Value = True Then
            If InStr(1, LCase(strCond), "where") > 0 Then
                strCond = strCond & " And (LicenseExpDate='' Or LicenseExpDate is null) "
            Else
                strCond = strCond & " Where (LicenseExpDate='' Or LicenseExpDate is null) "
            End If
        End If
        If Option3.Value = True Then
            If InStr(1, LCase(strCond), "where") > 0 Then
                strCond = strCond & " AND (AcceptedDate >= '" & LastYear & "') "
                strCond = strCond & " AND (AcceptedDate <= '" & OneMonth & "') "
            Else
                strCond = strCond & " WHERE (AcceptedDate >= '" & LastYear & "') "
                strCond = strCond & " AND (AcceptedDate <= '" & OneMonth & "') "
            End If
        End If
        If Option4.Value = True Then
            If InStr(1, LCase(strCond), "where") > 0 Then
                strCond = strCond & " AND (LicenseExpDate >= '" & TodayDate & "') "
                strCond = strCond & " AND (LicenseExpDate <= '" & CurYear & "') "
            Else
                strCond = strCond & " WHERE (LicenseExpDate >= '" & TodayDate & "') "
                strCond = strCond & " AND (LicenseExpDate <= '" & CurYear & "') "
            End If
        End If
    End If
    curCondition_ = strCond
    Screen.MousePointer = 11
    lvResult.ListItems.Clear
    For i = 1 To 1
        Set tmpRec = New ADODB.Recordset
        If i = 1 Then
'            If isEnv = True Then
                strCond = strCond & " And (Cstatus = 'N' Or Cstatus = 'P')"
'                strCond = strCond & " And (Cstatus = 'N' Or Cstatus = 'P') and companyid='776EDA5830F744A6B3551348D851BC22' "
'                Else
'                strCond = strCond & " And companyid='776EDA5830F744A6B3551348D851BC22' "
'            End If
            Sql1 = sqlStr & sqlStr1 & strCond
            Set tmpRec = GetRS(sqlStr & sqlStr1 & strCond & " Order By WasteNo", adOpenForwardOnly, adLockReadOnly)
'            tmpRec.Open sqlStr & sqlStr1 & strCond & " Order By WasteNo", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
'        ElseIf i = 2 Then
'            If InStr(1, LCase(strCond), "where") > 0 Then
'                strCond = strCond & " And (Status <> 'P') And  isAcceptedByEnvi <> 'Y'"
'            Else
'                strCond = strCond & " Where (Status <> 'P') "
'            End If
'            If isEnv = True Then
'                strCond = strCond & " And (Cstatus = 'N' Or Cstatus = 'P') "
'            End If
'            Sql2 = sqlStr & sqlStr2 & strCond
'            Set tmpRec = GetRS(sqlStr & sqlStr2 & strCond & " Order By WasteNo", adOpenForwardOnly, adLockReadOnly)
'            tmpRec.Open sqlStr & sqlStr2 & strCond & " Order By WasteNo", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
        End If
        If tmpRec.EOF = False Then
            TotalCount = TotalCount + tmpRec.RecordCount
            tmpArr = tmpRec.GetRows
            For x = 0 To UBound(tmpArr, 2)
                If Mid(tmpArr(19, x), 1, 1) = "Y" Then
                    Reject = True
                Else
                    Reject = False
                End If
                Set tmpList = lvResult.ListItems.Add(, "'" & tmpArr(0, x) & "'", FormatYMD_to_DMY("" & tmpArr(1, x), "/", "/"))
                statusDesc = ""
                If tmpArr(25, x) = "Y" Then
                    If Trim(tmpArr(24, x)) = "P" Then
                        statusDesc = "ขอออกใบเสนอราคาแบบเร่งด่วน ตรวจรับไปแล้ว ยังไม่สรุปผล"
                        curColor = lblApproved1.ForeColor
                    ElseIf Trim(tmpArr(24, x)) = "A" Then
                        curColor = lbApproved.ForeColor
                        statusDesc = "ขอออกใบเสนอราคาแบบเร่งด่วน ตรวจรับไปแล้ว รับกำจัด"
                    ElseIf Trim(tmpArr(24, x)) = "E" Then
                        curColor = lbClosed.ForeColor
                        statusDesc = "ขอออกใบเสนอราคาแบบเร่งด่วน ตรวจรับไปแล้ว ไม่รับกำจัด"
                    ElseIf Trim(tmpArr(24, x)) = "W" Then
                        curColor = lblQApprove.ForeColor
                        statusDesc = "รออนุมัติใบเสนอราคา (Waste ที่ฃอเสนอราคาแบบเร่งด่วนให้ Sales เก็บข้อมูลเพิ่ม)"
                    ElseIf Trim(tmpArr(24, x)) = "R" Then
                        curColor = lblQApprove.ForeColor
                        statusDesc = "รอตรวจรับ (Waste ที่ขอเสนอราคาแบบเร่งด่วน)"
                    ElseIf Trim(tmpArr(24, x)) = "J" Then
                        curColor = lblQApprove.ForeColor
                        statusDesc = "ตีกลับ ห้ามออกใบเสนอราคา (แบบเร่งด่วน)"
                    Else
                        statusDesc = "ขอออกใบเสนอราคาแบบเร่งด่วน"
                        curColor = vbBlack
                    End If
                ElseIf Trim(tmpArr(24, x)) = "J" Then
                    statusDesc = "ตีกลับ"
                    curColor = lblReject.ForeColor
                ElseIf Trim(tmpArr(24, x)) = "RJ" Then
                    statusDesc = "ตีกลับ ต้องการ Revised"
                    curColor = lblReject.ForeColor
                ElseIf Trim(tmpArr(24, x)) = "D" Then
                    statusDesc = "ยังไม่ส่งข้อมูล"
                    curColor = vbBlack
                ElseIf Trim(tmpArr(24, x)) = "R" Then
                    statusDesc = "รอตรวจรับ"
                    curColor = lbOpen1.ForeColor
                ElseIf Trim(tmpArr(24, x)) = "P" Then
                    statusDesc = "ยังไม่สรุปผล"
                    curColor = lblApproved1.ForeColor
                ElseIf Trim(tmpArr(24, x)) = "A" Then
                    curColor = lbApproved.ForeColor
                    statusDesc = "รับกำจัด"
                ElseIf Trim(tmpArr(24, x)) = "E" Then
                    curColor = lbClosed.ForeColor
                    statusDesc = "ไม่รับกำจัด"
                ElseIf Trim(tmpArr(24, x)) = "AR" Then
                    curColor = &H800080
                    statusDesc = "ต้องการ Revised Waste"
                ElseIf Trim(tmpArr(24, x)) = "IA" Then
                    curColor = &H800080
                    statusDesc = " Waste ถูก Revised"
                ElseIf tmpArr(12, x) = "Y" Then
                    curColor = lbCancel.ForeColor
                    statusDesc = "ยกเลิก"
                End If
                Dim AppLab As String
                  If tmpArr(37, x) = "D" Then
                        AppLab = "ยังไม่ได้ดำเนินการใดๆ"
                    ElseIf tmpArr(37, x) = "S" Then
                        AppLab = "ส่งวิเคราะห์"
                    ElseIf tmpArr(37, x) = "R" Then
                        AppLab = "รอวิเคราะห์"
                    ElseIf tmpArr(37, x) = "A" Then
                        AppLab = "วิเคราะห์ เสร็จเรียบร้อย"
                    Else
                        AppLab = ""
                    End If
                    Dim LapApprove As String
                    If tmpArr(38, x) = "Y" Then
                       LapApprove = "อนุมัติผลวิเคราะห์แล้ว"
                    ElseIf tmpArr(38, x) = "N" Then
                       LapApprove = "ไม่อนุมัติผลวิเคราะห์"
                    Else
                        LapApprove = ""
                    End If
                    Dim NewOld As String
                    If tmpArr(39, x) <> "" Then
                        If UCase(tmpArr(39, x)) <> "NULL" Then
                           NewOld = "Waste เก่า"
                           Else
                           NewOld = "Waste ใหม่"
                           NewCount = NewCount + 1
                        End If
                    Else
                       NewOld = "Waste ใหม่"
                        NewCount = NewCount + 1
                    End If
                    
                    
                If tmpArr(31, x) = "J" Then
                     statusDesc = "ลูกค้าโดนตีกลับ"
                     curColor = lblReject.ForeColor
                End If
                tmpList.ForeColor = curColor
                tmpList.SubItems(1) = IIf(IsNull(tmpArr(2, x)), "", Format(tmpArr(2, x), "dd/mm/yyyy"))
                tmpList.SubItems(2) = IIf(IsNull(tmpArr(4, x)), "", tmpArr(4, x))
                tmpList.SubItems(3) = IIf(IsNull(tmpArr(5, x)), "", tmpArr(5, x))
                tmpList.SubItems(4) = Trim("" & tmpArr(6, x)) & "-" & Trim("" & tmpArr(7, x)) & "-" & Trim("" & tmpArr(8, x))
                tmpList.SubItems(5) = statusDesc
                tmpList.SubItems(6) = IIf(IsNull(tmpArr(13, x)), "", Format(tmpArr(13, x), "dd/mm/yyyy"))
                tmpList.SubItems(7) = IIf(IsNull(tmpArr(14, x)), "", tmpArr(14, x))
                tmpList.SubItems(8) = IIf(IsNull(tmpArr(15, x)), "", tmpArr(15, x))
                tmpList.SubItems(9) = IIf(IsNull(tmpArr(16, x)), "", tmpArr(16, x))
                tmpList.SubItems(10) = IIf(IsNull(tmpArr(19, x)), "", tmpArr(19, x))
                tmpList.SubItems(11) = IIf(IsNull(tmpArr(20, x)), "", tmpArr(20, x))
                tmpList.SubItems(12) = IIf(IsNull(tmpArr(21, x)), "", tmpArr(21, x))
                tmpList.SubItems(13) = IIf(IsNull(tmpArr(22, x)), "", tmpArr(22, x))
                tmpList.SubItems(14) = IIf(IsNull(tmpArr(23, x)), "", tmpArr(23, x))
                tmpList.SubItems(15) = IIf(IsNull(tmpArr(35, x)), "", Format(tmpArr(35, x), "dd/mm/yyyy"))
                tmpList.SubItems(16) = IIf(IsNull(tmpArr(36, x)), "", Format(tmpArr(36, x), "dd/mm/yyyy"))
                tmpList.SubItems(17) = AppLab
                tmpList.SubItems(18) = LapApprove
                tmpList.SubItems(19) = NewOld
                For Y = 1 To tmpList.ListSubItems.Count
                    tmpList.ListSubItems(Y).ForeColor = curColor
                Next
                Set tmpList = Nothing
            Next
            Erase tmpArr
        End If
    Next i
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(TotalCount, "#,##0") & " รายการ    พบข้อมูล Waste ใหม่ " & Format(NewCount, "#,##0") & " รายการ"
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub
