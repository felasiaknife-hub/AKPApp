VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmMenifestClose 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ปิดใบ Menifest - F058"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelClosed 
      Caption         =   "ยกเลิกการปิดงาน"
      Height          =   795
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3300
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "อนุมัติการปิดงาน"
      Height          =   795
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3285
      Width           =   1485
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3300
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ปิดงาน"
      Height          =   795
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3300
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ข้อมูลเพิ่มเติมในการปิด Menifest"
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   270
      TabIndex        =   20
      Top             =   1020
      Width           =   8445
      Begin VB.TextBox txtCloseDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1755
      End
      Begin VB.TextBox txtM4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1530
         Width           =   465
      End
      Begin VB.TextBox txtH4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1530
         Width           =   465
      End
      Begin VB.TextBox txtM3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   1140
         Width           =   465
      End
      Begin VB.TextBox txtH3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "00"
         Top             =   1140
         Width           =   465
      End
      Begin VB.TextBox txtM2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   750
         Width           =   465
      End
      Begin VB.TextBox txtH2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   750
         Width           =   465
      End
      Begin VB.TextBox txtM1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "00"
         Top             =   360
         Width           =   465
      End
      Begin VB.TextBox txtH1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "00"
         Top             =   360
         Width           =   465
      End
      Begin DateCtl.Date dtDate1 
         Height          =   315
         Left            =   5670
         TabIndex        =   2
         Top             =   360
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   16777215
      End
      Begin DateCtl.Date dtDate2 
         Height          =   315
         Left            =   5670
         TabIndex        =   5
         Top             =   750
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtDate3 
         Height          =   315
         Left            =   5670
         TabIndex        =   8
         Top             =   1140
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปิด Menifest"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   4410
         TabIndex        =   32
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   2820
         TabIndex        =   31
         Top             =   1590
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เวลาที่กลับถึง Site งาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   630
         TabIndex        =   30
         Top             =   1590
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่กลับถึง Site งาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4170
         TabIndex        =   29
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   2820
         TabIndex        =   28
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เวลาที่ออกจากโรงงานลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   330
         TabIndex        =   27
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ออกจากโรงงานลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   3870
         TabIndex        =   26
         Top             =   810
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   2820
         TabIndex        =   25
         Top             =   810
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เวลาที่ถึงโรงงานลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   690
         TabIndex        =   24
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ถึงโรงงานลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   4200
         TabIndex        =   23
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2820
         TabIndex        =   22
         Top             =   420
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เวลาที่รถออกจาก Site งาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   390
         TabIndex        =   21
         Top             =   420
         Width           =   1815
      End
   End
   Begin VB.TextBox txtMenifestNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   540
      Width           =   2565
   End
   Begin VB.TextBox txtWorkDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Width           =   1755
   End
   Begin VB.TextBox txtCustName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   150
      Width           =   4995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manifest No."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   19
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6510
      TabIndex        =   18
      Top             =   210
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   720
      TabIndex        =   17
      Top             =   210
      Width           =   375
   End
End
Attribute VB_Name = "frmMenifestClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curMenifest_ As BWGMenifest
Dim wasteRec_ As BWGWasteReciep
Private Sub ApproveMenifest()
    Dim ResultStr As String
    Dim ResultStr2 As String
    ResultStr = curMenifest_.ApproveMenifest
    If ResultStr <> "" Then MsgBox ResultStr
    frmMenifest.SetCancelScreen
End Sub

Private Sub cmdApprove_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmMenifestClose.Name, ApproveRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    If curMenifest_.isClosed = True Then
        If curMenifest_.isApproved = False Then
            If MsgBox("ยืนยันการอนุมัติการปิดงาน ชื่อของท่านจะถูกบันทึกเป็นผู้อนุมัติการปิดงาน", vbExclamation + vbYesNo) = vbNo Then Exit Sub
            curMenifest_.isApproved = True
            Set wasteRec_ = New BWGWasteReciep
            wasteRec_.menifestID = curMenifest_.ID
             wasteRec_.Save
            ApproveMenifest
        Else
            If MsgBox("ใบเมนิเฟสนี้ ถูกอนุมัติการปิดงานแล้ว โดย " & curMenifest_.ApproveStaffName & _
                "ท่านต้องการยกเลิกการอนุมัติใช่หรือไม่", vbExclamation + vbYesNo) = vbNo Then Exit Sub
            curMenifest_.isApproved = False
            ApproveMenifest
            Call SetCaptionButton
        End If
    Else
        MsgBox "ใบเมนิเฟสนี้ยังไม่ได้ปิดงาน ไม่สามารถอนุมัติการปิดงานได้", vbInformation
    End If
End Sub

Private Sub cmdCancelClosed_Click()
    Dim ResultStr As String
    If curMenifest_.isApproved = True Then
        MsgBox "ใบเมนิเฟสนี้ถูกอนุมัติแล้วไม่สามารถยกเลิกได้ ถ้าต้องการยกเลิกการปิดงาน ให้ยกเลิการอนุมัติก่อน", vbInformation
        Exit Sub
    Else
        If MsgBox("ยืนยันการยกเลิกการปิดงาน", vbInformation + vbYesNo) = vbNo Then Exit Sub
        curMenifest_.isClosed = False
        ResultStr = curMenifest_.CloseMenifest
        If ResultStr <> "" Then MsgBox ResultStr
        Call SetCaptionButton
        frmMenifest.SetCancelScreen
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    SaveData True
End Sub

Public Function SaveData(ByVal showMsg As Boolean) As String
On Error Resume Next
Dim tmpRec As New ADODB.Recordset
Dim x%, sqlStr As String
Dim ErrMsg As String
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanUpdate = False Then
        MsgBox "ท่านไม่มีสิทธิบันทึกข้อมูล", vbExclamation, "Warning"
        Exit Function
    End If

    If Trim(txtH1.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtH1.SetFocus
        Exit Function
    End If
    If Trim(txtM1.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtM1.SetFocus
        Exit Function
    End If
    If Trim(txtH2.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtH2.SetFocus
        Exit Function
    End If
    If Trim(txtM2.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtM2.SetFocus
        Exit Function
    End If
    If Trim(txtH3.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtH3.SetFocus
        Exit Function
    End If
    If Trim(txtM3.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtM3.SetFocus
        Exit Function
    End If
    If Trim(txtH4.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtH4.SetFocus
        Exit Function
    End If
    If Trim(txtM4.Text) = "" Then
        MsgBox "กรุณาระบุเวลาต่างๆ ให้ครบถ้วนก่อนทำการเก็บข้อมูล", vbExclamation
        txtM4.SetFocus
        Exit Function
    End If
    If showMsg = True Then
        If MsgBox("ท่านต้องการปิดใบ Manifest นี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            With curMenifest_
                .DateArriveCust = dtDate1.ValueYMD
                .DateOutCust = dtDate2.ValueYMD
                .DateArriveSite = dtDate3.ValueYMD
                '.Save '*** ปิดฟังก์ชั่น Save ใช้ Update แทน เพราะคิดว่าไม่จำเป็นต้องบันทึกอะไรซ้ำอีก
                ErrMsg = .Update("UPDATE tbMenifestHeader SET DateArriveCust = '" & dtDate1.ValueYMD & "', DateOutCust = '" & .DateOutCust & "', DateArriveSite = '" & .DateArriveCust & "' WHERE MenifestID = '" & .ID & "'")
                curMenifest_.isClosed = True
                ErrMsg = ErrMsg & .CloseMenifest
                Call SetCaptionButton
                frmMenifest.SetCancelScreen
            End With
            
            'อัพเดทตาราง ServiceOrderHD สำหรับโปรแกรม AKPCustomer
            sqlStr = "SELECT * FROM tbJobDataHeader WHERE JobID='" & curMenifest_.CreatedFromJobID & "' and (Service_OrderID is not null and Service_OrderID<>'') "
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
                If tmpRec.RecordCount > 0 Then
                    DBConnExc "Update Service_OrderHD Set Completed='Y' WHERE OrderID='" & tmpRec("Service_OrderID") & "'"
                    tmpRec.Close
                    Set tmpRec = Nothing
                End If
        End If
    Else
        With curMenifest_
            .DateArriveCust = dtDate1.ValueYMD
            .DateOutCust = dtDate2.ValueYMD
            .DateArriveSite = dtDate3.ValueYMD
            '.Save '*** ปิดฟังก์ชั่น Save ใช้ Update แทน เพราะคิดว่าไม่จำเป็นต้องบันทึกอะไรซ้ำอีก
            ErrMsg = .Update("UPDATE tbMenifestHeader SET DateArriveCust = '" & dtDate1.ValueYMD & "', DateOutCust = '" & .DateOutCust & "', DateArriveSite = '" & .DateArriveCust & "' WHERE MenifestID = '" & .ID & "'")
            ErrMsg = ErrMsg & .CloseMenifest
        End With
    End If
    
    SaveData = Trim(ErrMsg & Trim(err.Description))
    cmdSave.Enabled = Not curMenifest_.isClosed
End Function

Private Sub dtDate1_Validate(Cancel As Boolean)
    curMenifest_.DateArriveCust = dtDate1.ValueYMD
End Sub

Private Sub dtDate2_Validate(Cancel As Boolean)
    curMenifest_.DateOutCust = dtDate2.ValueYMD
End Sub

Private Sub dtDate3_Validate(Cancel As Boolean)
    curMenifest_.DateArriveSite = dtDate2.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    txtCloseDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    dtDate1.Enabled = False
    dtDate2.Enabled = False
    dtDate3.Enabled = True
    txtH4.Text = Hour(Time)
    txtM4.Text = Minute(Time)
End Sub

Private Sub SetCaptionButton()
    If curMenifest_.isApproved = False Then
        cmdApprove.Caption = "อนุมัติการปิดงาน"
    Else
        cmdApprove.Caption = "ยกเลิกการอนุมัติ"
    End If
    
    If curMenifest_.isClosed = True Then
        cmdCancelClosed.Visible = True
        cmdApprove.Enabled = True
    Else
        cmdSave.Enabled = True
        cmdCancelClosed.Visible = False
        cmdApprove.Enabled = False
    End If
End Sub

Public Sub SetMenifestData(selMenifest As BWGMenifest)
    Set curMenifest_ = selMenifest
    With curMenifest_
        Call SetCaptionButton
        txtCustName.Text = .CustomerName
        txtWorkDate.Text = FormatYMD_to_DMY(.WorkDate, "/", "/")
        txtMenifestNo.Text = .MenifestNo
        If InStr(1, Trim(.TimeOutSite), ":") > 0 Then
            txtH1.Text = Trim(Split(.TimeOutSite, ":")(0))
            txtM1.Text = Trim(Split(.TimeOutSite, ":")(1))
        End If
        If InStr(1, Trim(.TimeArriveCust), ":") > 0 Then
            txtH2.Text = Trim(Split(.TimeArriveCust, ":")(0))
            txtM2.Text = Trim(Split(.TimeArriveCust, ":")(1))
        End If
        If InStr(1, Trim(.TimeOutCust), ":") > 0 Then
            txtH3.Text = Trim(Split(.TimeOutCust, ":")(0))
            txtM3.Text = Trim(Split(.TimeOutCust, ":")(1))
        End If
        If InStr(1, Trim(.TimeArriveSite), ":") > 0 Then
            txtH4.Text = Trim(Split(.TimeArriveSite, ":")(0))
            txtM4.Text = Trim(Split(.TimeArriveSite, ":")(1))
        End If
        If .isClosed = False Then
            dtDate1.ValueYMD = FormatDMY_to_YMD(txtWorkDate.Text, "/", "/")
            dtDate2.ValueYMD = FormatDMY_to_YMD(txtWorkDate.Text, "/", "/")
            dtDate3.ValueYMD = FormatDMY_to_YMD(txtWorkDate.Text, "/", "/")
        Else
            dtDate1.ValueYMD = .DateArriveCust
            dtDate2.ValueYMD = .DateOutCust
            dtDate3.ValueYMD = .DateArriveSite
        End If
        If IsDate(.CloseDate) Then txtCloseDate.Text = FormatYMD_to_DMY(.CloseDate, "/", "/")
        cmdSave.Enabled = Not .isClosed
    End With
End Sub

Private Sub txtH1_Validate(Cancel As Boolean)
    curMenifest_.TimeOutSite = Trim(txtH1.Text) & ":" & Trim(txtM1.Text)
End Sub

Private Sub txtH2_Validate(Cancel As Boolean)
    curMenifest_.TimeArriveCust = Trim(txtH2.Text) & ":" & Trim(txtM2.Text)
End Sub

Private Sub txtH3_Validate(Cancel As Boolean)
    curMenifest_.TimeOutCust = Trim(txtH3.Text) & ":" & Trim(txtM3.Text)
End Sub

Private Sub txtH4_Validate(Cancel As Boolean)
    curMenifest_.TimeArriveSite = Trim(txtH4.Text) & ":" & Trim(txtM4.Text)
End Sub

Private Sub txtM1_Validate(Cancel As Boolean)
    curMenifest_.TimeOutSite = Trim(txtH1.Text) & ":" & Trim(txtM1.Text)
End Sub

Private Sub txtM2_Validate(Cancel As Boolean)
    curMenifest_.TimeArriveCust = Trim(txtH2.Text) & ":" & Trim(txtM2.Text)
End Sub

Private Sub txtM3_Validate(Cancel As Boolean)
    curMenifest_.TimeOutCust = Trim(txtH3.Text) & ":" & Trim(txtM3.Text)
End Sub

Private Sub txtM4_Validate(Cancel As Boolean)
    curMenifest_.TimeArriveSite = Trim(txtH4.Text) & ":" & Trim(txtM4.Text)
End Sub
