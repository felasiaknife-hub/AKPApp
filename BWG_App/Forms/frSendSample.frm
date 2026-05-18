VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendSample 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F119"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12330
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox isApproveQT 
      Height          =   255
      Left            =   2880
      TabIndex        =   39
      Top             =   8400
      Width           =   210
   End
   Begin VB.CommandButton cmdQuo 
      Caption         =   "ใบเสนอราคา"
      Height          =   795
      Left            =   1560
      TabIndex        =   36
      Top             =   8340
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์ข้อมูล"
      Height          =   795
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8340
      Width           =   945
   End
   Begin VB.TextBox txtSendBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   10470
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   330
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดใบส่งตัวอย่าง"
      Height          =   2895
      Left            =   30
      TabIndex        =   20
      Top             =   750
      Width           =   9135
      Begin VB.ComboBox cboLabfrom2 
         Height          =   315
         Left            =   1740
         TabIndex        =   34
         Top             =   1800
         Width           =   3525
      End
      Begin VB.TextBox txtprice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   32
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cboLabfrom 
         Height          =   315
         Left            =   1740
         TabIndex        =   30
         Top             =   1440
         Width           =   3525
      End
      Begin VB.CheckBox chkIncTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ส่ง LAB ภายนอก"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1710
         TabIndex        =   6
         Top             =   2490
         Width           =   1815
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   660
         Width           =   7335
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3525
      End
      Begin VB.TextBox txtLABVender 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2430
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.TextBox txtIssuedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2595
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   240
         Width           =   4125
      End
      Begin WasteManagment.ctlDate dtWorkDate 
         Height          =   315
         Left            =   7185
         TabIndex        =   2
         Top             =   240
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาวิเคราห์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   5470
         TabIndex        =   35
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "รูปแบบวิเคราะห์ (2)"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   33
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "รูปแบบวิเคราะห์ (1)"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   31
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LAB Vender"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3540
         TabIndex        =   27
         Top             =   2490
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ที่อยู่"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1350
         TabIndex        =   26
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ส่งตัวอย่าง"
         Height          =   195
         Index           =   0
         Left            =   6090
         TabIndex        =   24
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ส่งตัวอย่าง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5580
         TabIndex        =   23
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   1290
         TabIndex        =   22
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "บริษัทที่ส่ง"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   21
         Top             =   1140
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "รายการของเสีย"
      Height          =   795
      Left            =   30
      TabIndex        =   17
      Top             =   8340
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8340
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8340
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8340
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8340
      Width           =   945
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "รายการตัวอย่าง"
      Height          =   4305
      Left            =   30
      TabIndex        =   18
      Top             =   3690
      Width           =   12255
      Begin VB.TextBox txtRem 
         Height          =   735
         Left            =   1200
         TabIndex        =   37
         Top             =   2400
         Width           =   7335
      End
      Begin VB.TextBox txtRemXXX 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   1560
         TabIndex        =   10
         Top             =   7340
         Width           =   1145
      End
      Begin WasteManagment.ctlGrid grSendSample 
         Height          =   1905
         Left            =   90
         TabIndex        =   8
         Top             =   330
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   4419
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   3885
         Left            =   8610
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   6853
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "พารามิเตอร์"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "หมายเหตุ"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "พารามิเตอร์"
         Height          =   195
         Index           =   2
         Left            =   8610
         TabIndex        =   29
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   4980
         Width           =   675
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8340
      Width           =   945
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "อนุมัติใช้ใบเสนอราคา"
      Height          =   255
      Left            =   3240
      TabIndex        =   40
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบส่งตัวอย่าง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   9120
      TabIndex        =   25
      Top             =   390
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบส่งตัวอย่างกากของเสีย"
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
      Left            =   150
      TabIndex        =   19
      Top             =   90
      Width           =   2145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   7650
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   13105
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frSendSample.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmSendSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colLABParam_ As Collection
Dim curSaleStaff_ As BWGCompanyStaff
Dim mySendSample_ As BWGSendSampleHeader
Dim curSendSampleDet_ As BWGSendSampleDetail
Dim curCustomer_ As BWGCustomer
Dim curLABVender_ As BWGTransportVendor
Dim colWasteDataCR_ As Collection
Dim myTemplate_ As BWGLABTemplateHeader
Dim curLab_ As Collection
Dim tmpRec As ADODB.Recordset
Dim isDetailChange_ As Boolean

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            mySendSample_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).id
        Else
            mySendSample_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
        End If
    Else
        mySendSample_.CompanyID = ""
    End If
End Sub

Private Sub cboLabfrom_Validate(Cancel As Boolean)
  If cboLabfrom.ListIndex >= 0 Then
        If cboLabfrom.ListCount > 1 Then
            If cboLabfrom.ListIndex = 0 Then Exit Sub
               mySendSample_.TemplateID = curLab_(cboLabfrom.ListIndex).id
        Else
               mySendSample_.TemplateID = curLab_(cboLabfrom.ListIndex).id
        End If
        txtPrice.Text = curLab_(cboLabfrom.ListIndex).Cost
        mySendSample_.CostLab = curLab_(cboLabfrom.ListIndex).Cost
    Else
        mySendSample_.TemplateID = ""
    End If
End Sub

Private Sub cboLabfrom2_Validate(Cancel As Boolean)
If cboLabfrom2.ListIndex >= 0 Then
        If cboLabfrom2.ListCount > 1 Then
            If cboLabfrom2.ListIndex = 0 Then Exit Sub
               mySendSample_.TemplateID2 = curLab_(cboLabfrom2.ListIndex).id
        Else
               mySendSample_.TemplateID2 = curLab_(cboLabfrom2.ListIndex).id
        End If
    Else
        mySendSample_.TemplateID2 = ""
    End If
End Sub

Private Sub chkIncTrans_Validate(Cancel As Boolean)
    If chkIncTrans.Value = Checked Then
        mySendSample_.IsTestExtByVender = True
    Else
        mySendSample_.IsTestExtByVender = False
    End If
End Sub

Private Sub cmdAddItem_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าก่อนทำการเพิ่มรายการกากของเสีย", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmWasteDataAdd
    With frmWasteDataAdd
        .SetCallerForm Me, curCustomer_
        .Show 1
        If isDetailChange_ Then
            ShowSendDetailData
            'ShowSendSampleDetails
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = mySendSample_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            txtCustomer.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub
Private Sub SetEnabledScreen()
Dim eFlag As Boolean
'    eFlag = Not mySendSample_.Delete
'    cmdApproved.Enabled = eFlag
'    cmdAddItem.Enabled = eFlag
'    cmdSave.Enabled = eFlag
'    cmdCancel.Enabled = eFlag
'    Command1.Enabled = eFlag
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If mySendSample_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewSendSample
    dtWorkDate.SetFocus
End Sub

Private Sub ClearScreen()
Dim x%
    txtSendBillNo.Text = "== AUTO =="
    dtWorkDate.ValueYMD = TodayDate
    txtCustomer.Text = ""
    txtAddr.Text = ""
    Set curCustomer_ = Nothing
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Set curSaleStaff_ = Nothing
    txtRem.Text = ""
    txtLABVender.Text = ""
    chkIncTrans.Value = 0
    grSendSample.ClearAllData
    grSendSample.getGridObj.Rows = 50
    
    For x = 1 To lvList.ListItems.Count
        lvList.ListItems(x).Checked = False
    Next
End Sub

Private Sub cmdPrint_Click()
Dim tmpStr As String, x%, totalPage%, tmpArr
Dim tmpCol As Collection, tmpDet As BWGSendSampleDetail
Dim pageCount%, rowCount%, y%, tmpLabParam As BWGLABParameterList
Dim avaiParams As String, z%, paramArr
Dim selResult As Integer
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนพิมพ์หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If mySendSample_.isChanged Then Exit Sub
        Else
            Exit Sub
        End If
    End If

    avaiParams = "ph,cu,cr,cd,ba,mn,ni,pb,zn,fe,as,se,hg,ag"
    paramArr = Split(avaiParams, ",")
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\SampleForm.rpt"
        .Formulas(0) = "CustName=" '& Chr(34) & Trim(txtCustomer.Text) & Chr(34)
        .Formulas(1) = "CustAddr=" & Chr(34) & Trim(txtAddr.Text) & Chr(34)
        .Formulas(2) = "SampleCount=" & mySendSample_.SendSampleDetList.Count
        .Formulas(3) = "ShowSendDate=" & Chr(34) & dtWorkDate.ValueDMY & Chr(34)
        
        Set tmpCol = mySendSample_.SendSampleDetList
        tmpArr = Split(Format(tmpCol.Count / 4, "0.0"), ".")
        If Trim(tmpArr(1)) <> "0" Then
            totalPage = CInt(Trim(tmpArr(0))) + 1
        Else
            totalPage = CInt(Trim(tmpArr(0)))
        End If
        .Formulas(4) = "ShowTotalPage=" & totalPage
        pageCount = 1
        rowCount = 1
        For x = 1 To tmpCol.Count
            .Formulas(5) = "ShowPageNo=" & pageCount
            Set tmpDet = tmpCol(x)
            
            .Formulas(20 + (rowCount - 1) * 6) = "ItemNo" & rowCount & "=" & x
            .Formulas(21 + (rowCount - 1) * 6) = "ItemName" & rowCount & "=" & Chr(34) & tmpDet.WasteDataName & Chr(34)
            .Formulas(22 + (rowCount - 1) * 6) = "ItemCode" & rowCount & "=" & Chr(34) & tmpDet.WasteNo & Chr(34)
            If tmpDet.SampleType = "SL" Then 'ของแข็ง
                .Formulas(23 + (rowCount - 1) * 6) = "WType" & rowCount & "=" & Chr(34) & "S" & Chr(34)
            ElseIf tmpDet.SampleType = "LQ" Then 'ของแหลว
                .Formulas(23 + (rowCount - 1) * 6) = "WType" & rowCount & "=" & Chr(34) & "L" & Chr(34)
            End If
            If tmpDet.PackageType = "BT" Then
                .Formulas(24 + (rowCount - 1) * 6) = "BTAmt" & rowCount & "=" & tmpDet.SampleQty
                .Formulas(25 + (rowCount - 1) * 6) = "BagAmt" & rowCount & "=0"
            ElseIf tmpDet.PackageType = "BG" Then
                .Formulas(24 + (rowCount - 1) * 6) = "BTAmt" & rowCount & "=0"
                .Formulas(25 + (rowCount - 1) * 6) = "BagAmt" & rowCount & "=" & tmpDet.SampleQty
            End If
            
            .Formulas(101 + (rowCount - 1) * 15) = "Param" & rowCount & "=" & Chr(34) & "" & Chr(34)
            '.Formulas(102 + (rowCount - 1) * 15) = "Param" & rowCount & "=" & Chr(34) & "" & Chr(34)
            '.Formulas(103 + (rowCount - 1) * 15) = "Param" & rowCount & "=" & Chr(34) & "" & Chr(34)
            '.Formulas(104 + (rowCount - 1) * 15) = "Param" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(105 + (rowCount - 1) * 15) = "Ba" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(106 + (rowCount - 1) * 15) = "Mn" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(107 + (rowCount - 1) * 15) = "Ni" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(108 + (rowCount - 1) * 15) = "Pb" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(109 + (rowCount - 1) * 15) = "Zn" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(110 + (rowCount - 1) * 15) = "Fe" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(111 + (rowCount - 1) * 15) = "As" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(112 + (rowCount - 1) * 15) = "Se" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(113 + (rowCount - 1) * 15) = "Hg" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(114 + (rowCount - 1) * 15) = "Ag" & rowCount & "=" & Chr(34) & "" & Chr(34)
'            .Formulas(115 + (rowCount - 1) * 15) = "Other" & rowCount & "=" & Chr(34) & "" & Chr(34)
            tmpStr = ""
            For y = 1 To tmpDet.SampleParamReq.Count
                Set tmpLabParam = tmpDet.SampleParamReq(y)
'                If InStr(1, avaiParams, LCase(tmpLabParam.ParameterCode)) Then
'                    For z = 0 To UBound(paramArr)
'                        If Trim(paramArr(z)) = LCase(tmpLabParam.ParameterCode) Then
'                            .Formulas(101 + z + (rowCount - 1) * 15) = Trim(tmpLabParam.ParameterCode) & rowCount & "=" & Chr(34) & "P" & Chr(34)
'                            Exit For
'                        End If
'                    Next z
'                End If
                tmpStr = tmpStr & tmpLabParam.ParameterName & ", "
                Set tmpLabParam = Nothing
            Next y
            If Trim(tmpStr) <> "" Then
                .Formulas(101 + (rowCount - 1) * 15) = "Param" & rowCount & "=" & Chr(34) & Left(tmpStr, Len(Trim(tmpStr)) - 1) & Chr(34)
            End If
            rowCount = rowCount + 1
            Set tmpDet = Nothing
        
            If x Mod 4 = 0 Then
'                For y = rowCount To 4
'                    .Formulas(20 + (y - 1) * 6) = "ItemNo" & y & "=0"
'                    .Formulas(21 + (y - 1) * 6) = "ItemName" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(22 + (y - 1) * 6) = "ItemCode" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(23 + (y - 1) * 6) = "WType" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(24 + (y - 1) * 6) = "BTAmt" & y & "=0"
'                    .Formulas(25 + (y - 1) * 6) = "BagAmt" & y & "=0"
'                    .Formulas(101 + (y - 1) * 15) = "pH" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(102 + (y - 1) * 15) = "Cu" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(103 + (y - 1) * 15) = "Cr" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(104 + (y - 1) * 15) = "Cd" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(105 + (y - 1) * 15) = "Ba" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(106 + (y - 1) * 15) = "Mn" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(107 + (y - 1) * 15) = "Ni" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(108 + (y - 1) * 15) = "Pb" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(109 + (y - 1) * 15) = "Zn" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(110 + (y - 1) * 15) = "Fe" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(111 + (y - 1) * 15) = "As" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(112 + (y - 1) * 15) = "Se" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(113 + (y - 1) * 15) = "Hg" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(114 + (y - 1) * 15) = "Ag" & y & "=" & Chr(34) & "" & Chr(34)
'                    .Formulas(115 + (y - 1) * 15) = "Other" & y & "=" & Chr(34) & "" & Chr(34)
'                Next y
                
                .Destination = crptToWindow
                .PrintReport
                If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
                pageCount = pageCount + 1
                rowCount = 1
            End If
        Next x
        If mySendSample_.SendSampleDetList.Count Mod 4 <> 0 Then
            For x = rowCount To 4
                .Formulas(20 + (x - 1) * 6) = "ItemNo" & x & "=0"
                .Formulas(21 + (x - 1) * 6) = "ItemName" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(22 + (x - 1) * 6) = "ItemCode" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(23 + (x - 1) * 6) = "WType" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(24 + (x - 1) * 6) = "BTAmt" & x & "=0"
                .Formulas(25 + (x - 1) * 6) = "BagAmt" & x & "=0"
                .Formulas(101 + (x - 1) * 15) = "pH" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(102 + (x - 1) * 15) = "Cu" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(103 + (x - 1) * 15) = "Cr" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(104 + (x - 1) * 15) = "Cd" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(105 + (x - 1) * 15) = "Ba" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(106 + (x - 1) * 15) = "Mn" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(107 + (x - 1) * 15) = "Ni" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(108 + (x - 1) * 15) = "Pb" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(109 + (x - 1) * 15) = "Zn" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(110 + (x - 1) * 15) = "Fe" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(111 + (x - 1) * 15) = "As" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(112 + (x - 1) * 15) = "Se" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(113 + (x - 1) * 15) = "Hg" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(114 + (x - 1) * 15) = "Ag" & x & "=" & Chr(34) & "" & Chr(34)
                .Formulas(115 + (x - 1) * 15) = "Other" & x & "=" & Chr(34) & "" & Chr(34)
            Next x
            
            .Destination = crptToWindow
            .PrintReport
            If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
        End If
    End With
End Sub

Private Sub cmdQuo_Click()
Dim tmpStr As String, x%, totalPage%, tmpArr
Dim tmpCol As Collection, tmpDet As BWGSendSampleDetail
Dim pageCount%, rowCount%, y%, tmpLabParam As BWGLABParameterList
Dim avaiParams As String, z%, paramArr
Dim selResult As Integer
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนพิมพ์หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If mySendSample_.isChanged Then Exit Sub
        Else
            Exit Sub
        End If
    End If

    avaiParams = "ph,cu,cr,cd,ba,mn,ni,pb,zn,fe,as,se,hg,ag"
    paramArr = Split(avaiParams, ",")
    Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If curCustomer_.CustomerType = "C" Or curCustomer_.CustomerType = "S" Or curCustomer_.CustomerType = "K" Or curCustomer_.CustomerType = "E" Or curCustomer_.CustomerType = "G" Or curCustomer_.CustomerType = "T" Then
            
            If curCustomer_.CustomerType = "K" Or curCustomer_.CustomerType = "T" Then
                .ReportFileName = App.Path & "\reports\QuotationLabTK.rpt"
            ElseIf curCustomer_.CustomerType = "G" Or curCustomer_.CustomerType = "E" Then
                  .ReportFileName = App.Path & "\reports\QuotationLabGE.rpt"
            Else
'                If mySendSample_.CompanyID = "C451AF07C9A44B4282D6E08B14AC0195" Then
'                    ReportFileName = App.Path & "\reports\QuotationLabBWC.rpt"
'                Else
                    .ReportFileName = App.Path & "\reports\QuotationLabCs.rpt"
'                End If
            End If
        Else
            .ReportFileName = App.Path & "\reports\QuotationLab.rpt"
        End If
        tmpSql = "select * from vw_SendSampleParameter  " & vbCrLf
'        tmpSql = tmpSql & " where SendSampleDetID = '0BCA9C044C2B4B698F1F5D17B1E0A22E' and WasteDataID = '3601B608EB0D4E909E95D656AF4CFCA7' " & vbCrLf
        tmpSql = tmpSql & " where SendSampleID = '" & mySendSample_.id & "' " & vbCrLf
        tmpSql = tmpSql & " order by ParameterName " & vbCrLf
        Dim CusName  As String
        Dim cusAddr As String
        If curCustomer_.CustomerType = "P" Or curCustomer_.CustomerType = "S" Or curCustomer_.CustomerType = "K" Or curCustomer_.CustomerType = "E" Then
            Dim sStr As String
            sStr = "select Addr1 ,TumbolName ,AumphurName ,ProvinceName ,ZipCode from vw_CustomerAddress Where CustomerID = '" & curCustomer_.ParentCustID & "' "
            Set tmpRec = GetRS(sStr, adOpenForwardOnly, adLockReadOnly)
            If tmpRec.EOF = False Then
            cusAddr = tmpRec(0) & "  ต." & tmpRec(1) & "  อ." & tmpRec(2) & "  จ." & tmpRec(3) & " " & tmpRec(4)
            End If
            CusName = txtCustomer.Text & "(" & curCustomer_.ParentCustName & ")"
        Else
             CusName = txtCustomer.Text
             cusAddr = txtAddr.Text
        End If
        .Formulas(0) = "CustomerName=" & Chr(34) & CusName & Chr(34)
        .Formulas(1) = "Address=" & Chr(34) & cusAddr & Chr(34)
        .Formulas(2) = "WasteName=" & Chr(34) & grSendSample.getGridObj.TextMatrix(1, 1) & Chr(34)
        .Formulas(3) = "Cost=" & Chr(34) & Format(txtPrice.Text, "#,##0.00") & Chr(34)
        Dim QNo As String
        QNo = "QT-" & txtSendBillNo.Text
        .Formulas(4) = "SendNo=" & Chr(34) & QNo & Chr(34)
        .Formulas(5) = "QDate=" & Chr(34) & dtWorkDate.ValueDMY & Chr(34)
        .Formulas(6) = "Tel=" & Chr(34) & curCustomer_.TelNo & Chr(34)
        .Formulas(7) = "Fax=" & Chr(34) & curCustomer_.FaxNo & Chr(34)
        .Formulas(8) = "CustName=" & Chr(34) & curCustomer_.ContactPerson & Chr(34)
        Dim salename As String
        salename = curCustomer_.UnderSaleName & "  " & curCustomer_.UnderSaleTelNo
        .Formulas(9) = "ContakeName=" & Chr(34) & salename & Chr(34)
        
        .WindowTitle = "ใบเสนอราคาวิเคราะห์ "
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql
        .Destination = 0
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    
End Sub

Private Sub CmdSave_Click()
Dim x%
Dim tmpDet As BWGSendSampleDetail
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ส่ง", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If

    If chkIncTrans.Value = Checked Then
        If Trim(txtLABVender.Text) = "" Then
            MsgBox "กรุณาระบุLAB Vender", vbExclamation
            txtLABVender.SetFocus
            Exit Sub
        End If
    End If
    If txtPrice.Text = "" Then
        mySendSample_.CostLab = "0"
    End If
    
    If mySendSample_.isChanged Then
        mySendSample_.Save
        txtSendBillNo.Text = mySendSample_.SendBillNo
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGSendSampleDetail
Dim tmpCol As Collection
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If mySendSample_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmSendSampleSearch
    With frmSendSampleSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Public Sub SetDataSendSample(selSend As BWGSendSampleHeader)
Dim tmpStr
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim tmpColLab1 As Collection
Dim tmpColLab2 As Collection

    Set mySendSample_ = selSend
    With mySendSample_
        txtSendBillNo.Text = .SendBillNo
        dtWorkDate.ValueYMD = .SendDate
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        'txtAddr.Text = .CustomerName
        
        Set tmpCol = tmpS.CompanySearch("CompanyID='" & .CompanyID & "'")
        Set tmpColLab1 = tmpS.LABTemplateHeaderSearch("TemplateID='" & .TemplateID & "'")
        Set tmpColLab2 = tmpS.LABTemplateHeaderSearch("TemplateID='" & .TemplateID2 & "'")
        
        If tmpCol.Count > 0 Then
            SetTextToCombo tmpCol(1).CompanyNameTH, cboCompany
        End If
        If tmpColLab1.Count > 0 Then
            SetTextToCombo tmpColLab1(1).TemplateName, cboLabfrom
        Else
            cboLabfrom.ListIndex = 0
        End If
        If tmpColLab2.Count > 0 Then
            SetTextToCombo tmpColLab2(1).TemplateName, cboLabfrom2
        Else
            cboLabfrom2.ListIndex = 0
        End If
        txtIssuedStaff.Text = .IssuedStaffName
        txtPrice.Text = .CostLab
        If .IsTestExtByVender Then
            chkIncTrans.Value = 1
        Else
            chkIncTrans.Value = 0
        End If
        If .isApproveQT = "Y" Then
            isApproveQT.Value = 1
        Else
            isApproveQT.Value = 0
        End If
        Set tmpCol = tmpS.TransportVenderSearch("VendorID='" & .LABVenderID & "'")
        If tmpCol.Count > 0 Then SetDataVenderLAB tmpCol(1), False
        'txtLABVender = .LABVenderName
        txtRem.Text = .Remdesc
        ShowSendDetailData
    End With
    Set tmpCol = Nothing
End Sub

Private Sub ShowSendDetailData()
Dim x%, tmpDet As BWGSendSampleDetail
    With grSendSample.getGridObj
        grSendSample.ClearAllData
        .Rows = mySendSample_.SendSampleDetList.Count + 50
        For x = 1 To mySendSample_.SendSampleDetList.Count
            Set tmpDet = mySendSample_.SendSampleDetList(x)
            .TextMatrix(x, 0) = tmpDet.id
            .TextMatrix(x, 1) = tmpDet.WasteDataName
            .TextMatrix(x, 2) = tmpDet.WasteNo
            If tmpDet.SampleType = "SL" Then
                .TextMatrix(x, 3) = "ของแข็ง"
            ElseIf tmpDet.SampleType = "LQ" Then
                .TextMatrix(x, 3) = "ของเหลว"
            End If
            If tmpDet.PackageType = "BT" Then
                .TextMatrix(x, 4) = "ขวด"
            ElseIf tmpDet.PackageType = "BG" Then
                .TextMatrix(x, 4) = "ถุง"
            End If
            .TextMatrix(x, 5) = tmpDet.SampleQty
            Set tmpDet = Nothing
        Next
    End With
End Sub
Private Sub dtWorkDate_Validate(Cancel As Boolean)
    mySendSample_.SendDate = dtWorkDate.ValueYMD
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mySendSample_ = Nothing
Set curCustomer_ = Nothing
Set curLABVender_ = Nothing
End Sub

Private Sub grSendSample_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 5 Then
       If NewValue <> "" Then
        If IsNumeric(NewValue) = False Then
            MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
            Cancel = True
        End If
       End If
    End If
End Sub

Private Sub grSendSample_Click(row As Integer, col As Integer)
    'MsgBox grSendSample.getGridObj.ColWidth(col)
End Sub

Private Sub grSendSample_GotFocus()
    If grSendSample.getGridObj.col = 0 Then grSendSample.getGridObj.col = 1
End Sub

Private Sub grSendSample_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grSendSample.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการทดสอบนี้หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = mySendSample_.SendSampleDetList
                selCol.Remove tmpKey
                curSendSampleDet_.isChanged = True
                mySendSample_.isChanged = True
            End If
        End If
    End With
End Sub

Private Sub grSendSample_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGSendSampleDetail
Dim tmpS As New BWGSearchManager
Dim tmpKey As String
Dim tmpVal As String
Dim tmpRow%, x%

    tmpKey = Trim(grSendSample.getGridObj.TextMatrix(row, 0))
    If col = 1 And Trim(grSendSample.getGridObj.TextMatrix(row, col)) = "" Then Exit Sub
    If isExist(mySendSample_.SendSampleDetList, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGSendSampleDetail
        tmpDet.id = GetGUID
        tmpDet.SendSampleID = mySendSample_.id
        tmpDet.RowPos = row
        grSendSample.getGridObj.TextMatrix(row, 0) = tmpDet.id
        mySendSample_.SendSampleDetList.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = mySendSample_.SendSampleDetList("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grSendSample.getGridObj.TextMatrix(row, col))

    With tmpDet
        Select Case col
            Case 1
                .WasteDataName = Trim(tmpVal)
            Case 2
                '.WasteDataID = Trim(tmpVal)
            Case 3
                If Trim(tmpVal) <> "" Then
                   If Trim(tmpVal) = "ของแข็ง" Then
                        .SampleType = "SL"
                   ElseIf Trim(tmpVal) = "ของเหลว" Then
                        .SampleType = "LQ"
                   End If
                End If
            Case 4
                If Trim(tmpVal) <> "" Then
                   If Trim(tmpVal) = "ขวด" Then
                        .PackageType = "BT"
                   ElseIf Trim(tmpVal) = "ถุง" Then
                        .PackageType = "BG"
                   End If
                End If
            Case 5
                If Trim(tmpVal) <> "" Then
                    .SampleQty = tmpVal
                End If
        End Select

        mySendSample_.isChanged = True
    End With
    Set tmpDet = Nothing

End Sub


Private Sub isApproveQT_Validate(Cancel As Boolean)
    If isApproveQT.Value = 1 Then
        mySendSample_.isApproveQT = "Y"
    Else
        mySendSample_.isApproveQT = "N"
    End If
End Sub

Private Sub lvList_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If curSendSampleDet_ Is Nothing Then
        MsgBox "กรุณาเลือกรายการของเสีย"
        Exit Sub
    End If
    If curSendSampleDet_.id = "" Then
        MsgBox "กรุณาเลือกรายการของเสีย"
        Exit Sub
    End If
    If item.Checked Then
        If isExist(curSendSampleDet_.SampleParamReq, item.Key) = False Then
            curSendSampleDet_.SampleParamReq.Add colLABParam_(item.Key), item.Key
        End If
    Else
        If isExist(curSendSampleDet_.SampleParamReq, item.Key) Then
            curSendSampleDet_.SampleParamReq.Remove item.Key
        End If
    End If
    curSendSampleDet_.isChanged = True
    mySendSample_.isChanged = True
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With mySendSample_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With mySendSample_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If mySendSample_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
        txtAddr.Text = .Address + " " + .TumbolName + " " + .AumphurName + " " + .ProvinceName + " " + .ZipCode
        If isAssigned Then
            mySendSample_.CustomerID = .id
            mySendSample_.CustomerName = .CustomerName
            Set mySendSample_.SendSampleDetList = Nothing
            grSendSample.ClearAllData
            grSendSample.getGridObj.Rows = 50
        End If
    End With
End Sub

Private Sub chkIncTrans_Click()
    If chkIncTrans.Value = Checked Then
        Label2.item(4).Visible = True
        txtLABVender.Visible = True
        'txtLABVender.SetFocus
    Else
        Label2.item(4).Visible = False
        txtLABVender.Visible = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub ShowLABParamList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGLABParameterList
Dim searchM As New BWGSearchManager
    
    Set colLABParam_ = searchM.ParameterLABSearch
    
    If colLABParam_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colLABParam_.Count
        Set tmpItem = colLABParam_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.id & "'", tmpItem.ParameterName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub SetMSFlexGrid()
Dim x%
    grSendSample.ColInputType(3) = ComboBoxEnum
    grSendSample.ColInputType(4) = ComboBoxEnum
   
    With grSendSample.getGridObj
        .Rows = 50
        .Cols = 6
        .TextMatrix(0, 1) = "ชื่อตัวอย่าง"
        .TextMatrix(0, 2) = "รหัส"
        .TextMatrix(0, 3) = "ประเภท"
        .TextMatrix(0, 4) = "ภาชนะ"
        .TextMatrix(0, 5) = "จำนวน"
        .ColWidth(0) = 0
        .ColWidth(1) = 3600
        .ColWidth(2) = 1305
        .ColWidth(3) = 1230
        .ColWidth(4) = 1080
        .ColWidth(5) = 1005
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
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
        cmdPrint.Picture = .picPrint.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    
    ShowLABParamList
    PopulateCompany cboCompany, cboIndex
    PopulateLabFromMe
    dtWorkDate.ValueYMD = TodayDate
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    CreateNewSendSample
    
End Sub
Private Sub PopulateLabFromMe()
Dim x%, tmpS As New BWGSearchManager
    cboLabfrom.Clear
    cboLabfrom2.Clear
    cboLabfrom.AddItem "==เลือก=="
    cboLabfrom2.AddItem "==เลือก=="
     Set curLab_ = tmpS.LabFromSearch
    For x = 1 To curLab_.Count
        cboLabfrom.AddItem curLab_(x).TemplateName
        cboLabfrom2.AddItem curLab_(x).TemplateName
    Next
    If cboLabfrom.ListCount > 0 Then
        cboLabfrom.ListIndex = 0
    End If
    If cboLabfrom2.ListCount > 0 Then
        cboLabfrom2.ListIndex = 0
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grSendSample" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub grSendSample_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpS As BWGSearchManager
Dim tmpCbo As ComboBox
Dim tmpKey As String
Dim tmpCol As Collection
Dim y%
    
    Set tmpCol = mySendSample_.SendSampleDetList

    With grSendSample.getGridObj
        If Trim(.TextMatrix(row - 1, 1)) = "" And col > 1 Then
            Label1(2).Caption = "พารามิเตอร์"
            grSendSample.ColEnabled(col) = False
            Exit Sub
        Else
            If Trim(.TextMatrix(row, 1)) <> "" Then
                Label1(2).Caption = "พารามิเตอร์สำหรับ " & .TextMatrix(row, 1)
            Else
                Label1(2).Caption = "พารามิเตอร์"
            End If
           If col = 3 Or col = 4 Or col = 5 Then
                grSendSample.ColEnabled(col) = True

            Else
                grSendSample.ColEnabled(col) = False
            End If
        End If
    End With
        For x = 1 To lvList.ListItems.Count
            lvList.ListItems(x).Checked = False
        Next
        tmpKey = "'" & grSendSample.getGridObj.TextMatrix(row, 0) & "'"
        If isExist(mySendSample_.SendSampleDetList, tmpKey) Then
            Set curSendSampleDet_ = mySendSample_.SendSampleDetList(tmpKey)
            For x = 1 To curSendSampleDet_.SampleParamReq.Count
                If isExist(lvList.ListItems, "'" & curSendSampleDet_.SampleParamReq(x).id & "'") Then
                    lvList.ListItems("'" & curSendSampleDet_.SampleParamReq(x).id & "'").Checked = True
                End If
            Next
        End If
        If (col = 3) Then
            Set tmpCbo = grSendSample.getColObject(3)
            tmpCbo.Clear
            tmpCbo.AddItem "ของแข็ง"
            tmpCbo.AddItem "ของเหลว"
            Set tmpCbo = Nothing
        ElseIf (col = 4) Then
            Set tmpCbo = grSendSample.getColObject(4)
            tmpCbo.Clear
            tmpCbo.AddItem "ขวด"
            tmpCbo.AddItem "ถุง"
            Set tmpCbo = Nothing
        End If

End Sub

Private Sub CreateNewSendSample()
    Set mySendSample_ = Nothing
    Set mySendSample_ = New BWGSendSampleHeader
    Set curSendSampleDet_ = Nothing
    Set curSendSampleDet_ = New BWGSendSampleDetail
    
    With mySendSample_
        .id = GetGUID
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
            End If
        End If
        .SendDate = TodayDate
        .IssuedStaffID = CurrentUser.id
        .SendStaffID = CurrentUser.id
        .SampleRecStaffID = CurrentUser.id
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .IsTestExtByVender = False
        .HeaderType = "A"
        .isChanged = False
    End With
    
    With curSendSampleDet_
        .id = GetGUID
        .SendSampleID = mySendSample_.id
        .isChanged = False
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If mySendSample_ Is Nothing Then Exit Sub
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            'cmdSave_Click
            If mySendSample_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtLABVender_Validate(Cancel As Boolean)
    If curLABVender_ Is Nothing Then Set curLABVender_ = New BWGTransportVendor
    If Trim(txtLABVender.Text) = "" And Trim(curLABVender_.NameTH) <> Trim(txtLABVender.Text) Then
        Set curLABVender_ = Nothing
        Set curLABVender_ = New BWGTransportVendor
        With mySendSample_
            .LABVenderID = ""
            .LABVenderName = ""
        End With
    Else
        If Trim(curLABVender_.NameTH) <> Trim(txtLABVender.Text) Then
            Set curLABVender_ = Nothing
            Set curLABVender_ = New BWGTransportVendor
            With mySendSample_
                .LABVenderID = ""
                .LABVenderName = ""
            End With
            Load frmLABVenderSearch
            With frmLABVenderSearch
                .getCriteria (txtLABVender.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If mySendSample_.LABVenderID = "" Then txtLABVender.Text = ""
        End If
    End If
End Sub

Public Sub SetDataVenderLAB(selTransportVender As BWGTransportVendor, Optional isAssigned As Boolean = True)
Dim x%
    Set curLABVender_ = Nothing
    Set curLABVender_ = selTransportVender
    If curLABVender_ Is Nothing Then Exit Sub
    With curLABVender_
        txtLABVender.Text = .NameTH
        If isAssigned Then
            mySendSample_.LABVenderID = .id
            mySendSample_.LABVenderName = .NameTH
        End If
    End With
End Sub




Private Sub txtprice_Validate(Cancel As Boolean)
    mySendSample_.CostLab = txtPrice.Text
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    mySendSample_.Remdesc = txtRem.Text
End Sub

Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGSendSampleDetail
    Set tmpCol = mySendSample_.SendSampleDetList

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.id)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).wastedataID = tmpID Then
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.id = GetGUID
        tmpDetail.SendSampleID = mySendSample_.id
        tmpDetail.wastedataID = selItem.id
        tmpDetail.WasteDataName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
        isDetailChange_ = True
        mySendSample_.isChanged = True
        Set tmpDetail = Nothing
    End If
    If mySendSample_.SendSampleDetList Is Nothing Then
        Set mySendSample_.SendSampleDetList = tmpCol
        isDetailChange_ = True
        mySendSample_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Private Sub ShowSendSampleDetails()
Dim x%, y%
Dim tmpCol As Collection
Dim tmpDet As BWGSendSampleDetail
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    With grSendSample.getGridObj
        grSendSample.ClearAllData
                
        Set tmpCol = mySendSample_.SendSampleDetList
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.id
            .row = x + rowCount
            .col = 1
            .CellForeColor = forColor
            .TextMatrix(x + rowCount, 1) = tmpDet.WasteDataName
            .TextMatrix(x + rowCount, 2) = tmpDet.WasteNo
            tmpDet.RowPos = x
        Next
        For x = 1 To curSendSampleDet_.SampleParamReq.Count
            lvList.ListItems("'" & curSendSampleDet_.SampleParamReq(x).id & "'").Checked = isExist(lvList.ListItems, "'" & curSendSampleDet_.SampleParamReq(x).id & "'")
        Next
'        For x = 1 To curSendSampleDet_.SampleParamReq.Count
'            If isExist(lvList.ListItems, "'" & curSendSampleDet_.SampleParamReq(x).ID & "'") Then
'                lvList.ListItems("'" & curSendSampleDet_.SampleParamReq(x).ID & "'").Checked = True
'            End If
'        Next
        .row = 1
        .col = 1
    End With
    Set tmpCol = Nothing
End Sub

Public Function getWastDataItems() As Collection
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGSendSampleDetail
Dim tmpWaste As BWGWasteDataCR
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = mySendSample_.SendSampleDetList
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.wastedataID) <> "" Then
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.id = tmpDet.wastedataID
                tmpWaste.WasteName = tmpDet.WasteDataName
                tmpWaste.WasteNo = tmpDet.WasteNo
                tmpCol.Add tmpWaste, "'" & tmpDet.wastedataID & "'"
                Set tmpWaste = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With

    Set getWastDataItems = tmpCol
    Set tmpCol = Nothing
End Function

Public Sub RemoveWasteData(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = mySendSample_.SendSampleDetList
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).wastedataID = selID Then
            tmpKey = "'" & tmpCol(x).id & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        mySendSample_.isChanged = True
    End If
End Sub

