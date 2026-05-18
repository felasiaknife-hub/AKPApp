VERSION 5.00
Begin VB.Form frmLABResult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F047"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์ผล LAB"
      Height          =   795
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8430
      Width           =   1185
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "First Approve"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8430
      Width           =   1185
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   30
      TabIndex        =   23
      Top             =   840
      Width           =   11325
      Begin VB.ComboBox cboRecType 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1380
         Width           =   4845
      End
      Begin VB.TextBox txtSampleDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   990
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2040
         Width           =   4845
      End
      Begin VB.TextBox txtWasteName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   990
         Width           =   4605
      End
      Begin VB.TextBox txtWasteNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   990
         Width           =   1665
      End
      Begin VB.TextBox txtAnalysisNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   2160
         Width           =   1755
      End
      Begin WasteManagment.ctlDate dtReport 
         Height          =   315
         Left            =   9180
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   6345
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   4845
      End
      Begin VB.TextBox txtReportNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   210
         Width           =   1755
      End
      Begin WasteManagment.ctlDate dtSampling 
         Height          =   315
         Left            =   9180
         TabIndex        =   7
         Top             =   990
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtReceived 
         Height          =   315
         Left            =   9180
         TabIndex        =   8
         Top             =   1380
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtAnalyticalEnd 
         Height          =   315
         Left            =   9180
         TabIndex        =   10
         Top             =   1770
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtAnalyticalStart 
         Height          =   315
         Left            =   7200
         TabIndex        =   9
         Top             =   1770
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Descriptoin"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   34
         Top             =   1770
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waste No"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   33
         Top             =   1050
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Analysis No."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   8250
         TabIndex        =   32
         Top             =   2220
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   9030
         TabIndex        =   31
         Top             =   1830
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Analytical Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   6030
         TabIndex        =   30
         Top             =   1830
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   8040
         TabIndex        =   29
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampling Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   8070
         TabIndex        =   28
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   8220
         TabIndex        =   27
         Top             =   660
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ที่อยู่"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   26
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   25
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report No."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   8310
         TabIndex        =   24
         Top             =   270
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameter LAB ที่ต้องการบันทึกผล"
      Height          =   4935
      Left            =   0
      TabIndex        =   22
      Top             =   3450
      Width           =   11355
      Begin WasteManagment.ctlGrid grLABDetail 
         Height          =   4605
         Left            =   60
         TabIndex        =   13
         Top             =   270
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   8123
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8430
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8430
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8430
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8430
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8430
      Width           =   1005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   5770
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   5890
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บันทึกผล LAB ของกากของเสีย"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   60
      Width           =   2685
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frLABResult.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmLABResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myLabResult_ As BWGLABResultHeader
Dim curCustomer_ As BWGCustomer
Dim curSampleDet_ As BWGSendSampleDetail
Dim curWaste_ As BWGWasteDataCR
Dim colLABMethod_ As Collection
Dim colLABParam_ As Collection

Private Sub SetMSFlexGrid()
    grLABDetail.ColInputType(4) = ComboBoxEnum
    grLABDetail.ColInputType(6) = ComboBoxEnum
    With grLABDetail.getGridObj
        .RowHeight(0) = 0
        .Cols = 9
        .ColWidth(0) = 0
        .ColWidth(1) = 1575
        .ColWidth(2) = 1380
        .ColWidth(3) = 3885
        .ColWidth(4) = 600
        .ColWidth(5) = 855
        .ColWidth(6) = 600
        .ColWidth(7) = 855
        .ColWidth(8) = 1230
        .Rows = 100
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 4
        .FixedAlignment(8) = 4
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 7
    
        .TextMatrix(0, 1) = "Parameter"
        .TextMatrix(0, 2) = "Unit"
        .TextMatrix(0, 3) = "Method Of Analysis..."
        .TextMatrix(0, 4) = "Sign Result1"
        .TextMatrix(0, 5) = "Result1"
        .TextMatrix(0, 6) = "Sign Result2"
        .TextMatrix(0, 7) = "Result2"
        .TextMatrix(0, 8) = "Standard"
    End With
End Sub
Private Sub cmdApprove_Click()
Dim x%
Dim tmpDet As BWGLABResultDetail
Dim selResult As Integer
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    For x = 1 To myLabResult_.LabResultDetails.Count
        Set tmpDet = myLabResult_.LabResultDetails(x)
        If Trim(tmpDet.MethodID) = "" Then
            MsgBox "กรุณาระบุข้อมูล Method ให้ครบถ้วน", vbExclamation
            Set tmpDet = Nothing
            Exit Sub
        End If
    Next
    Set tmpDet = Nothing
    If cboRecType.ListIndex = 0 Then
        MsgBox "กรุณาระบุการตรวจรับตัวอย่าง", vbExclamation
        cboRecType.SetFocus
        Exit Sub
    End If
    If cboRecType.ListIndex = 2 Then
        myLabResult_.AnalysisNo = "Rejected"
        myLabResult_.ReportNo = "Rejected"
        myLabResult_.isCancel = True
    End If
    If myLabResult_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLabResult_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    If MsgBox("ท่านต้องการอนุมัติผล LAB นี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        If myLabResult_.isApprove4Print = False Then
            myLabResult_.Approve4Print
            cmdApprove.Caption = "Final Approve"
        ElseIf myLabResult_.isFinalApproved = False Then
            myLabResult_.FinalApprove
            cmdApprove.Enabled = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลการบันทึกผล LAB นี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myLabResult_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            txtSampleDesc.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myLabResult_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLabResult_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    ClearScreen
    CreateNewLabREsult
    PopulateUnit
    PopulateLABMethod
    PopulateLABParam
    txtSampleDesc.SetFocus
End Sub
Private Sub ClearScreen()
    txtCustomer.Text = ""
    txtReportNo.Text = "== AUTO =="
    dtReport.ValueYMD = TodayDate
    dtSampling.ValueYMD = TodayDate
    dtReceived.ValueYMD = TodayDate
    dtAnalyticalStart.ValueYMD = TodayDate
    dtAnalyticalEnd.ValueYMD = TodayDate
    txtAnalysisNo.Text = "== AUTO =="
    txtSampleDesc.Text = ""
    txtAddr.Text = ""
    txtWasteNo.Text = ""
    txtWasteName.Text = ""
    grLABDetail.ClearAllData
    grLABDetail.getGridObj.Rows = 100
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If myLabResult_.isApprove4Print = False Then
        MsgBox "กรุณาทำ First Approve ก่อนสั่งพิมพ์", vbExclamation
        Exit Sub
    End If
End Sub

Private Sub CmdSave_Click()
Dim x%
Dim tmpDet As BWGLABResultDetail
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If

    For x = 1 To myLabResult_.LabResultDetails.Count
        Set tmpDet = myLabResult_.LabResultDetails(x)
        If Trim(tmpDet.MethodID) = "" Then
            MsgBox "กรุณาระบุข้อมูล Method ให้ครบถ้วน", vbExclamation
            Set tmpDet = Nothing
            Exit Sub
        End If
    Next
    Set tmpDet = Nothing
    If cboRecType.ListIndex = 0 Then
        MsgBox "กรุณาระบุการตรวจรับตัวอย่าง", vbExclamation
        cboRecType.SetFocus
        Exit Sub
    End If
    If cboRecType.ListIndex = 2 Then
        myLabResult_.AnalysisNo = "Rejected"
        myLabResult_.ReportNo = "Rejected"
        myLabResult_.isCancel = True
    End If
    If myLabResult_.isChanged Then
        myLabResult_.ReportNo = txtReportNo.Text
        myLabResult_.AnalysisNo = txtAnalysisNo.Text
        myLabResult_.Save
        txtReportNo.Text = myLabResult_.ReportNo
        txtAnalysisNo.Text = myLabResult_.AnalysisNo
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    
    If myLabResult_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLabResult_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmLabResultSearch
    With frmLabResultSearch
        .SetCallerForm Me
        .Show 1
    End With
    txtSampleDesc.SetFocus
End Sub

Private Sub dtAnalyticalEnd_Validate(Cancel As Boolean)
    myLabResult_.AnalyticalDateEnd = dtAnalyticalEnd.ValueYMD
End Sub

Private Sub dtAnalyticalStart_Validate(Cancel As Boolean)
    myLabResult_.AnalyticalDateStart = dtAnalyticalStart.ValueYMD
End Sub

Private Sub dtReceived_Validate(Cancel As Boolean)
    myLabResult_.ReceivedDate = dtReceived.ValueYMD
End Sub

Private Sub dtReport_Validate(Cancel As Boolean)
    myLabResult_.ReportDate = dtReport.ValueYMD
End Sub

Private Sub dtSampling_Validate(Cancel As Boolean)
    myLabResult_.SamplingDate = dtSampling.ValueYMD
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    Call SetMSFlexGrid
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboRecType
        .AddItem "== กรุณาระบุการตรวจรับตัวอย่างรายการนี้ =="
        .AddItem "สามารถรับวิเคราะห์ตัวอย่างนี้ได้"
        .AddItem "ไม่สามารถรับวิเคราะห์ตัวอย่างนี้"
        .ListIndex = 0
    End With
    Set myLabResult_ = Nothing
    Set myLabResult_ = New BWGLABResultHeader
     PopulateCompany cboCompany, cboIndex
     PopulateLABParam
     PopulateLABMethod
     PopulateUnit
     CreateNewLabREsult
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grLABDetail" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub CreateNewLabREsult()
Set myLabResult_ = Nothing
Set myLabResult_ = New BWGLABResultHeader
    With myLabResult_
        .ID = GetGUID
        If Not curCustomer_ Is Nothing Then
            .CustomerAddr = curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        If Not curWaste_ Is Nothing Then
           .wastedataID = curWaste_.ID
           .WasteName = curWaste_.WasteName
           .WasteNo = curWaste_.WasteNo
        End If
        .isCancel = False
        .ReportDate = dtReport.ValueYMD
        .SamplingDate = dtSampling.ValueYMD
        .ReceivedDate = dtReceived.ValueYMD
        .AnalyticalDateStart = dtAnalyticalStart.ValueYMD
        .AnalyticalDateEnd = dtAnalyticalEnd.ValueYMD
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                '.CompanyName = CurrentUser.WorkCompanies(cboCompany.ListIndex).CompanyNameTH
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                '.CompanyName = CurrentUser.WorkCompanies(cboCompany.ListIndex).CompanyNameTH
            End If
        End If
        .GetLabTemplete
        .isChanged = False
    End With
    cmdApprove.Caption = "First Approve"
    cmdApprove.Enabled = True
    SetEnabledScreen
    ShowLabResult
End Sub

Private Sub SetEnabledScreen()
Dim eFlag As Boolean
    eFlag = Not myLabResult_.isChanged
'    cmdApproved.Enabled = eFlag
'    cmdAddItem.Enabled = eFlag
    cmdSave.Enabled = eFlag
    cmdCancel.Enabled = eFlag
'    Command1.Enabled = eFlag
End Sub

Private Sub PopulateLABMethod()
Dim tmpS As New BWGSearchManager
    Set colLABMethod_ = tmpS.MethodLABSearch
    Set tmpS = Nothing
End Sub

Private Sub PopulateLABParam()
Dim tmpS As New BWGSearchManager
    Set colLABParam_ = tmpS.ParameterLABSearch
    Set tmpS = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myLabResult_ Is Nothing Then Exit Sub
    If myLabResult_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLabResult_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub grLABDetail_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
Dim tmpDet As BWGLABResultDetail
Dim tmpKey As String
    If Trim(grLABDetail.getGridObj.TextMatrix(row, 0)) = "" Then Exit Sub
    tmpKey = "'" & Trim(grLABDetail.getGridObj.TextMatrix(row, 0)) & "'"
    Set tmpDet = myLabResult_.LabResultDetails(tmpKey)
    If col = 5 And IsNumeric(NewValue) = False And tmpDet.ResultType = "N" Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลข", vbExclamation
        Cancel = True
    End If
    Set tmpDet = Nothing
End Sub

Private Sub grLABDetail_Click(row As Integer, col As Integer)
    'MsgBox grLABDetail.getGridObj.ColWidth(col)
End Sub

Private Sub grLABDetail_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
Dim tmpDet As BWGLABResultDetail
    With grLABDetail.getGridObj
        If Trim(Trim(.TextMatrix(.row, 0))) = "" Then Exit Sub
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            tmpKey = "'" & Trim(.TextMatrix(.row, 0)) & "'"
            Set tmpDet = myLabResult_.LabResultDetails(tmpKey)
            If .col = 4 Then
                tmpDet.CompareType = ""
            ElseIf .col = 5 Then
                tmpDet.LABResultNumber = 0
                tmpDet.LABResultString = ""
            End If
            .TextMatrix(.row, .col) = ""
            Set tmpDet = Nothing
            If grLABDetail.ColInputType(.col) = ComboBoxEnum Then grLABDetail.getColObject(.col).ListIndex = -1
            
        End If
    End With

End Sub

Private Sub grLABDetail_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpS As BWGSearchManager
Dim tmpCbo As ComboBox
Dim tmpKey As String
    
    With grLABDetail.getGridObj
        If Trim(.TextMatrix(row, 0)) = "" Then
            grLABDetail.ColEnabled(col) = False
        ElseIf col = 4 Or col = 5 Or col = 6 Or col = 7 Then
            grLABDetail.ColEnabled(col) = True
            If col = 4 Or col = 6 Then
                Set tmpCbo = grLABDetail.getColObject(4)
                tmpCbo.Clear
                tmpCbo.AddItem ">"
                tmpCbo.AddItem "<"
                tmpCbo.AddItem "="
                tmpCbo.AddItem "<="
                tmpCbo.AddItem ">="
                tmpCbo.AddItem "<>"
                tmpCbo.AddItem "N.A."
                tmpCbo.AddItem "N.P."
                Set tmpCbo = Nothing
            End If
        End If
    End With
End Sub

Private Sub grLABDetail_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGLABResultDetail
Dim tmpKey As String
Dim tmpVal As String
Dim tmpRow%

    tmpKey = Trim(grLABDetail.getGridObj.TextMatrix(row, 0))
    If Trim(tmpKey) = "" Then Exit Sub
    Set tmpDet = myLabResult_.LabResultDetails("'" & tmpKey & "'")
    tmpVal = Trim(grLABDetail.getGridObj.TextMatrix(row, col))

    With tmpDet
        Select Case col
            Case 4
'                If Trim(tmpVal) <> "" Then
'                    .LabUnitID = colUnit_(grLABDetail.getColObject(4).ListIndex + 1).ID
'                    .LabUnitName = Trim(tmpVal)
'                Else
'                    .LabUnitID = ""
'                    .LabUnitName = ""
'                End If
                .CompareType = grLABDetail.getColObject(4).Text
            Case 5
                If tmpDet.ResultType = "N" Then
                    .LABResultNumber = tmpVal
                ElseIf tmpDet.ResultType = "S" Then
                    .LABResultString = tmpVal
                End If
            Case 6
                .CompareType2 = grLABDetail.getColObject(6).Text
            Case 7
                If tmpDet.ResultType = "N" Then
                    .LABResultNumber2 = tmpVal
                ElseIf tmpDet.ResultType = "S" Then
                    .LABResultString2 = tmpVal
                End If
        End Select

        myLabResult_.isChanged = True
    End With
    Set tmpDet = Nothing

End Sub

Public Sub SetDataLabResult(selLab As BWGLABResultHeader)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myLabResult_ = selLab
    With myLabResult_
        txtReportNo.Text = .ReportNo
        dtReport.ValueYMD = .ReportDate
        dtSampling.ValueYMD = .SamplingDate
        dtReceived.ValueYMD = .ReceivedDate
        dtAnalyticalStart.ValueYMD = .AnalyticalDateStart
        dtAnalyticalEnd.ValueYMD = .AnalyticalDateEnd
        txtAnalysisNo.Text = .AnalysisNo
        txtSampleDesc.Text = .SampleDesc
        If Trim(.AnalysisNo) = "Rejected" Then
            cboRecType.ListIndex = 2
        Else
            cboRecType.ListIndex = 1
        End If
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        SetDataCustomer tmpCol(1), False
        Set tmpCol = tmpS.WasteDataCRSearch("WasteDataID='" & .wastedataID & "'")
        txtAnalysisNo.Text = "LB-" & Trim(.WasteNo)
        txtReportNo.Text = "RPT-" & Trim(.WasteNo)
        SetDataWasteData tmpCol(1), False
        cmdApprove.Enabled = True
        If .isApprove4Print Then
            If .isFinalApproved Then
                cmdApprove.Enabled = False
            Else
                cmdApprove.Caption = "Final Approve"
            End If
        Else
            cmdApprove.Caption = "First Approve"
        End If
    End With
    ShowLabResult
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
    Set curCustomer_ = selCust
    With curCustomer_
        txtCustomer.Text = .CustomerName
        txtAddr.Text = .Address & " " & .TumbolName & " " & .AumphurName & " " & .ProvinceName
        If isAssigned Then
            myLabResult_.CustomerAddr = txtAddr.Text
            myLabResult_.CustomerID = .ID
            myLabResult_.CustomerName = .CustomerName
        End If
    End With
End Sub

Public Sub SetDataWasteData(selWaste As BWGWasteDataCR, Optional isAssigned As Boolean = True)
    Set curWaste_ = selWaste
    With curWaste_
        txtWasteNo.Text = .WasteNo
        txtWasteName.Text = .WasteName
        If isAssigned Then
            myLabResult_.wastedataID = .ID
            myLabResult_.WasteName = .WasteName
            myLabResult_.WasteNo = .WasteNo
        End If
    End With
End Sub

Private Sub ShowLabResult()
Dim tmpDet As BWGLABResultDetail
Dim x%, curType As String, rowCount%, Y%
    With grLABDetail.getGridObj
        grLABDetail.ClearAllData
        .Rows = myLabResult_.LabResultDetails.Count + 1
        rowCount = 1
        For x = 1 To myLabResult_.LabResultDetails.Count
            Set tmpDet = myLabResult_.LabResultDetails(x)
            If Trim(curType) <> "Method of Analysis of " & Trim(tmpDet.TestTypeName) Then
                curType = "Method of Analysis of " & Trim(tmpDet.TestTypeName)
                For Y = 1 To .Cols - 1
                    .row = rowCount
                    .col = Y
                    .CellBackColor = &HC0C0C0
                    .CellFontBold = True
                    .CellAlignment = 4
                Next
                .Rows = .Rows + 1
                .TextMatrix(rowCount, 1) = "Parameter"
                .TextMatrix(rowCount, 2) = "Unit"
                .TextMatrix(rowCount, 3) = curType
                .TextMatrix(rowCount, 4) = "Sign1"
                .TextMatrix(rowCount, 5) = "Result1"
                .TextMatrix(rowCount, 6) = "Sign2"
                .TextMatrix(rowCount, 7) = "Result2"
                .TextMatrix(rowCount, 8) = "Standard"
                rowCount = rowCount + 1
            End If
            Set tmpDet = myLabResult_.LabResultDetails(x)
            .TextMatrix(rowCount, 0) = tmpDet.ID
            .TextMatrix(rowCount, 1) = tmpDet.ParameterName
            .TextMatrix(rowCount, 2) = tmpDet.LabUnitName
            .TextMatrix(rowCount, 3) = tmpDet.MethodName
            .TextMatrix(rowCount, 4) = tmpDet.CompareType
            If tmpDet.ResultType = "N" Then
                .TextMatrix(rowCount, 5) = tmpDet.LABResultNumber
                .TextMatrix(rowCount, 7) = tmpDet.LABResultNumber2
            Else
                .TextMatrix(rowCount, 5) = tmpDet.LABResultString
                .TextMatrix(rowCount, 7) = tmpDet.LABResultString2
            End If
            .TextMatrix(rowCount, 6) = tmpDet.CompareType2
            .TextMatrix(rowCount, 8) = tmpDet.StandardValue
            If Not curSampleDet_ Is Nothing Then
                If isExist(curSampleDet_.SampleParamReq, "'" & tmpDet.ParameterID & "'") Then
                    For Y = 1 To .Cols - 1
                        .row = rowCount
                        .col = Y
                        .CellForeColor = &HC00000
                        .CellFontBold = True
                    Next
                End If
            End If
            rowCount = rowCount + 1
        Next
        Set tmpDet = Nothing
    End With
End Sub

Private Sub txtSampleDesc_Validate(Cancel As Boolean)
    myLabResult_.SampleDesc = txtSampleDesc.Text
End Sub

Public Sub SetSampleDetID(selID As String)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set tmpCol = tmpS.SendSampleDetailSearch("SendSampleDetID='" & selID & "'")
    If tmpCol.Count > 0 Then Set curSampleDet_ = tmpCol(1)
    myLabResult_.SendSampleDetID = selID
    Set tmpCol = tmpS.LabResultHeaderSearch("SendSampleDetID='" & selID & "'")
    If tmpCol.Count > 0 Then
        SetDataLabResult tmpCol(1)
    Else
        Set tmpCol = tmpS.WasteDataCRSearch("WasteDataID in (Select WasteDataID from tbSendSampleDetail WHERE SendSampleDetID='" & selID & "')")
        If tmpCol.Count > 0 Then
            SetDataWasteData tmpCol(1)
            Set tmpCol = tmpS.CustomerSearch("CustomerID='" & curWaste_.CustomerID & "'")
            If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1)
        End If
    End If
    ShowLabResult
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub



