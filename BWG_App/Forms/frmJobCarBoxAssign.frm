VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobCarBoxAssign 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F039"
   ClientHeight    =   8835
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12405
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCloseJob 
      Caption         =   "ปิด Job"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7980
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdMenifest 
      Caption         =   "ออก Manifest"
      Height          =   795
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7980
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "ตารางการเดินรถ"
      Height          =   2115
      Left            =   60
      TabIndex        =   25
      Top             =   5760
      Width           =   12255
      Begin WasteManagment.ctlGrid grTimeTable 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   2990
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   7200
         Picture         =   "frmJobCarBoxAssign.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   315
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Left            =   60
         Top             =   240
         Width           =   12135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "รถที่ใช้ในการขน"
      Height          =   1335
      Left            =   60
      TabIndex        =   24
      Top             =   4320
      Width           =   12255
      Begin VSFlex7LCtl.VSFlexGrid grCar 
         Height          =   1320
         Left            =   60
         TabIndex        =   27
         ToolTipText     =   "ถ้าต้องการแก้ไขข้อความ ให้เลือกเซลล์แล้วกด F2 เพื่อแก้ไข"
         Top             =   240
         Width           =   12135
         _cx             =   21405
         _cy             =   2328
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   12
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   3525
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   3360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดใบเสนอราคา"
      Height          =   3555
      Left            =   60
      TabIndex        =   16
      Top             =   780
      Width           =   12255
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1740
         TabIndex        =   31
         Top             =   1800
         Width           =   3435
      End
      Begin VB.CheckBox ChkNotActive 
         BackColor       =   &H80000009&
         Caption         =   "งานขนส่งพิเศษ"
         Height          =   315
         Left            =   9180
         TabIndex        =   29
         Top             =   1380
         Width           =   1875
      End
      Begin DateCtl.Date ctlWorkDate 
         Height          =   315
         Left            =   4500
         TabIndex        =   28
         Top             =   300
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cboSiteName 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1020
         Width           =   9375
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   4
         Top             =   1380
         Width           =   7335
      End
      Begin VB.TextBox txtIssuedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   3195
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   660
         Width           =   9375
      End
      Begin VB.TextBox txtJobOpenNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   300
         Width           =   1755
      End
      Begin MSComctlLib.ListView lvResult 
         Height          =   1275
         Left            =   60
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2220
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   2249
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "หมายเหตุทั่วไปของลูกค้า"
            Object.Width           =   20814
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รูปแบบการวิ่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   30
         Top             =   1860
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   23
         Top             =   1440
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่เก็บขน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   660
         TabIndex        =   22
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ออกใบสั่งงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   6660
         TabIndex        =   20
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   660
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   18
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ขน"
         Height          =   195
         Index           =   8
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7980
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7980
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7980
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7980
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7980
      Width           =   945
   End
   Begin VB.Label lblLastUpdate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "แก้ไขล่าสุดโดย :"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8940
      TabIndex        =   26
      Top             =   0
      Width           =   3480
   End
   Begin VB.Label lbCancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   4680
      TabIndex        =   21
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
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
      Caption         =   "กำหนดข้อมูลรถให้กับใบงาน"
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
      TabIndex        =   15
      Top             =   90
      Width           =   2445
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
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frmJobCarBoxAssign.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmJobCarBoxAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myJobOpenBox_ As BWGJobDataBox
Dim curCustomer_ As BWGCustomer
Dim curSaleStaff_ As BWGCompanyStaff
Dim isDetailChange_ As Boolean
Dim curCar_ As BWGJobDataCarUsedBox
Dim curCarType_ As BWGTruckType

Dim TypeOfBox As String
Const TypeOfCar = "#1;พร้อม|#2;ไม่พร้อม"
Const TypeOfPhotography = "#1;ถ่าย|#2;ไม่ถ่าย"

Dim SelectBoxType As String
Dim SelectBoxCount As Integer

Dim TruckType As String

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myJobOpenBox_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            myJobOpenBox_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
    Else
        myJobOpenBox_.CompanyID = ""
    End If
End Sub

Private Sub cboSiteName_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If cboSiteName.ListIndex > -1 Then
        myJobOpenBox_.CustomerSiteID = curCustomer_.SiteAddresses(cboSiteName.ListIndex + 1).ID
        myJobOpenBox_.CustomerSiteName = cboSiteName.Text
    Else
        myJobOpenBox_.CustomerSiteID = ""
        myJobOpenBox_.CustomerSiteName = ""
    End If
End Sub



'Private Sub cboType_Validate(Cancel As Boolean)
'    If cboType.ListIndex = 0 Then 'เที่ยวหลัก
'        myJobOpenBox_.JobType = "A"
'    Else
'        myJobOpenBox_.JobType = "B"
'    End If
'End Sub

Private Sub cmdCancel_Click()
    Dim result
    result = MsgBox("ต้องการยกเลิกรายการนี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myJobOpenBox_.CancelJob
        If myJobOpenBox_.isCancel Then
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
        SetEnabledScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myJobOpenBox_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpenBox_.isChanged And cmdSave.Enabled Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewJobOpen
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



'Private Sub cmdMenifest_Click()
'Dim selResult As Integer
'Dim x%, y%, isCompleted As Boolean
'Dim itemCount%, tmpCar As BWGJobDataCarUsedBox, tmpTimeTable As BWGJobDataTimeTableBox
'    If Trim(txtCustomer.Text) = "" Then
'        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
'        txtCustomer.SetFocus
'        Exit Sub
'    End If
'    If myJobOpenBox_.JobDetails Is Nothing Then
'        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
'        Exit Sub
'    End If
'
''    For x = 1 To myJobOpenBox_.Cars.Count
''        Set tmpCar = myJobOpenBox_.Cars(x)
''        For y = 1 To tmpCar.TimeTables.Count
''            Set tmpTimeTable = tmpCar.TimeTables(y)
''            If Trim(tmpTimeTable.VendorTruckID) = "" Then
''                MsgBox "กรุณากำหนดรถให้ครบถ้วนก่อนทำการสร้างใบ Menifest", vbExclamation
''                Set tmpCar = Nothing
''                Set tmpTimeTable = Nothing
''                Exit Sub
''            End If
''            Set tmpTimeTable = Nothing
''        Next
''        Set tmpCar = Nothing
''    Next
'
'    isCompleted = myJobOpenBox_.isCompleted
'    If myJobOpenBox_.isChanged And cmdSave.Enabled Then
'        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
'        If selResult = vbYes Then
'            cmdSave_Click
'            If myJobOpenBox_.isChanged And cmdSave.Enabled Then Exit Sub
'        Else
'            MsgBox "ยังไม่ได้ทำการออกใบ Manifest", vbInformation
'            Exit Sub
'        End If
'    End If
'    Screen.MousePointer = 11
'    itemCount = myJobOpenBox_.GenerateMenifest
'    Screen.MousePointer = 0
'    MsgBox "จัดบ็อกสำเร็จ " & itemCount
'    With myJobOpenBox_
''        lbMenifestCount.Caption = .MenifestCountAll & "/" & .MenifestCountIssued
''        If isCompleted = False And .MenifestCountAll = .MenifestCountIssued Then
''            myJobOpenBox_.CloseJob
''        End If
'        cmdSave.Enabled = Not .MenifestCountAll = .MenifestCountIssued
'        cmdCancel.Enabled = cmdSave.Enabled
'        cmdMenifest.Enabled = cmdSave.Enabled
'        SetEnabledScreen
'    End With
'End Sub

Private Sub CmdSave_Click()
Dim x%, Y%, isCarInUse As Boolean
Dim tmpCol As Collection
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
'    If myJobOpenBox_.JobDetails Is Nothing Then
'        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
'        Exit Sub
'    End If
    
'    If myJobOpenBox_.JobDetails.Count = 0 Then
'        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
'        Exit Sub
'    End If
    If myJobOpenBox_.Cars Is Nothing Then
        MsgBox "กรุณาระบุรถที่ต้องใช้ในการขนอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpenBox_.Cars.Count = 0 Then
        MsgBox "กรุณาระบุรถที่ต้องใช้ในการขนอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    For x = 1 To myJobOpenBox_.Cars.Count
        If myJobOpenBox_.Cars(x).GoSameTime = "" Then
            MsgBox "กรุณาระบุว่ารถไปพร้อมกันหรือไม่ในแต่ละเที่ยวให้ครบถ้วน", vbExclamation
            Exit Sub
        End If
        For Y = 1 To myJobOpenBox_.Cars(x).TimeTables.Count
            If Trim(myJobOpenBox_.Cars(x).TimeTables(Y).StartTime) = "" Then
                MsgBox "กรุณาระบุตารางการเดินรถของรถคันที่ " & myJobOpenBox_.Cars(x).carNo & " ให้ครบถ้วน", vbExclamation
                Exit Sub
            End If
        Next
    Next
    If ChkNotActive.Value = 1 Then
        myJobOpenBox_.isNotActive = "Y"
    Else
        myJobOpenBox_.isNotActive = "N"
    End If
    If cboType.ListIndex = 0 Then
        myJobOpenBox_.JobType = "A"
    ElseIf cboType.ListIndex = 1 Then
        myJobOpenBox_.JobType = "B"
    ElseIf cboType.ListIndex = 2 Then
        myJobOpenBox_.JobType = "F"
    Else
        myJobOpenBox_.JobType = "N"
    End If
    
    myJobOpenBox_.Save
    
    If myJobOpenBox_.isChanged And myJobOpenBox_.isCancel = False And myJobOpenBox_.isCompleted = False Then
        Set tmpCol = myJobOpenBox_.Cars
        For x = 1 To myJobOpenBox_.Cars.Count
            Set tmpCol = myJobOpenBox_.Cars(x).TimeTables
            DBConnExc "Update tbJobDataCarUsedBox Set TruckSubTypeID='" & myJobOpenBox_.Cars(x).TruckSubTypeID & "',JobType = " & myJobOpenBox_.JobType & " ,isTrailer='" & TF_To_YN(myJobOpenBox_.Cars(x).isTrailer) & "' WHERE JobDataCarID='" & myJobOpenBox_.Cars(x).ID & "'"
'            DBConnExc "Update tbJobDataTimeTable Set JobType = " & myJobOpenBox_.JobType & "  WHERE JobDataCarID='" & myJobOpenBox_.Cars(X).ID & "'"
            For Y = 1 To tmpCol.Count
                If Trim(tmpCol(Y).VendorTruckID) = "" Then
                    MsgBox "ข้อมูลผิดพลาด กรุณาจดเลขที่ Job แล้วติดต่อผู้ดูแลระบบทันที", vbExclamation
                    Exit Sub
                Else
                    If Trim(tmpCol(Y).BoxType) <> "" And tmpCol(Y).BoxCount = 0 Then MsgBox "ไม่ได้กำหนดจำนวน box กรุณากำหนดจำนวน box", vbInformation
'                    DBConnExc "Update tbJobDataTimeTableBox Set VendorTruckID='" & tmpCol(y).VendorTruckID & "',DriverName='" & tmpCol(y).DriverName & "',TrailTruckID='" & tmpCol(y).TrailTruckID & "', BoxType = '" & tmpCol(y).BoxType & "', BoxCount=" & tmpCol(y).BoxCount & " WHERE JobDataCarID='" & tmpCol(y).JobDataCarID & "' And TripNo=" & tmpCol(y).TripNo
                    If Not tmpCol(Y) Is Nothing Then
                        tmpCol(Y).isChanged = True
                        tmpCol(Y).Save
                    End If
                End If
            Next Y
'            For y = 1 To tmpCol.Count
'                If Trim(tmpCol(y).TrailTruckID) <> "" Then
'                    DBConnExc "Update tbJobDataTimeTableBox Set TimeTableNo='" & tmpCol(y).TimeTableNo & "' WHERE JobDataCarID='" & tmpCol(y).JobDataCarID & "' And TripNo=" & (tmpCol(y).TripNo + 1) & " And VendorTruckID='" & tmpCol(y).TrailTruckID & "'"
'                End If
'            Next y
        
        Next x

        myJobOpenBox_.isChanged = False
    End If
    If myJobOpenBox_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpCol As Collection
    If myJobOpenBox_.isChanged And myJobOpenBox_.isCancel = False And myJobOpenBox_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpenBox_.isChanged And cmdSave.Enabled Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmJobOpenSearch
    With frmJobOpenSearch
        .JobType = "JobDataHeaderBoxSearch"
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdCloseJob_Click()
Dim result
    If myJobOpenBox_.isChanged And myJobOpenBox_.isCancel = False And cmdSave.Enabled Then
        result = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If result = vbYes Then
            CmdSave_Click
            If myJobOpenBox_.isChanged Then Exit Sub
        Else
            MsgBox "ยังไม่ได้ทำการปิด Job", vbExclamation
            Exit Sub
        End If
    End If
    result = MsgBox("ต้องการปิด Job นี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myJobOpenBox_.CloseJob
        SetEnabledScreen
    End If
End Sub

Private Sub ctlWorkDate_Validate(Cancel As Boolean)
    If Not myJobOpenBox_ Is Nothing Then
        myJobOpenBox_.WorkDate = ctlWorkDate.ValueYMD
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grWasteData" And Me.ActiveControl.Name <> "grCar" And Me.ActiveControl.Name <> "grTimeTable" Then
        SendKeys "{Tab}"
    ElseIf KeyCode = 27 Then
        cmdClose_Click
    End If
End Sub

Private Sub Form_Load()
Dim x%
'    CanAccess = False: CanUpdate = False
'    SetPermission Me, CanAccess, CanUpdate
'    If CanAccess = False Then
'        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
'        Unload Me
'    End If

    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    cboType.Clear
    With cboType
        .AddItem "เที่ยวหลัก"
        .AddItem "เที่ยวเสริม"
        .AddItem "งานส่ง RDF"
        .ListIndex = 0
    End With
    PopulateUnit
    PopulateCompany cboCompany, cboIndex
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    
    LoadBoxType
    Call SetGrid
    CreateNewJobOpen
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim MinWidth As Integer
Dim NewWidth As Integer, NewHeight As Integer
    If Me.Width < 12645 Then Me.Width = 12645
    If Me.Height < 8985 Then Me.Height = 9405

    NewWidth = (Me.Width - 12645)
    NewHeight = Me.Height - 9405
    MinWidth = 365

    cmdSave.Left = Me.Width - (cmdSave.Width * 5) - MinWidth - 50
    cmdClear.Left = Me.Width - (cmdClear.Width * 4) - MinWidth - 150
    cmdSearch.Left = Me.Width - (cmdSearch.Width * 3) - MinWidth - 100
    cmdCancel.Left = Me.Width - (cmdCancel.Width * 2) - MinWidth - 50
    cmdClose.Left = Me.Width - (cmdClose.Width * 1) - MinWidth


    cmdSave.Top = Me.Height - cmdSave.Height - 650
    cmdClear.Top = Me.Height - cmdClear.Height - 650
    cmdSearch.Top = Me.Height - cmdSearch.Height - 650
    cmdCancel.Top = Me.Height - cmdCancel.Height - 650
    cmdClose.Top = Me.Height - cmdClose.Height - 650

    Frame1.Width = Me.Width - MinWidth
    Frame3.Width = Me.Width - MinWidth
    Frame4.Width = Me.Width - MinWidth
    Frame4.Height = 2115 + NewHeight

    Shape2.Width = Frame4.Width - 150
    Shape2.Height = 1785 + NewHeight

    lvResult.Width = Frame1.Width - 200
    grCar.Width = Frame3.Width - 100
    grTimeTable.Width = Frame4.Width - 250
    grTimeTable.Height = 1725 + NewHeight - 50

    lvResult.ColumnHeaders(1).Width = lvResult.Width
    grCar.ColWidth(9) = 2610 + NewWidth
    grTimeTable.getGridObj.ColWidth(5) = 1800 + (NewWidth / 2)
    grTimeTable.getGridObj.ColWidth(grTimeTable.getGridObj.Cols - 1) = 1800 + (NewWidth / 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myJobOpenBox_ Is Nothing Then Exit Sub
    If myJobOpenBox_.isChanged And myJobOpenBox_.isCancel = False And cmdSave.Enabled And myJobOpenBox_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpenBox_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub SetGrid()
Dim x%
    With grCar
        .Rows = 20
        .Cols = 10
        .TextMatrix(0, 1) = "รถคันที่"
        .TextMatrix(0, 2) = "ประเภทรถ"
        .TextMatrix(0, 3) = "จำนวนเที่ยว"
        .TextMatrix(0, 4) = "คนงาน/คัน"
        .TextMatrix(0, 5) = "ไปพร้อมกัน?"
        .TextMatrix(0, 6) = "ถ่ายรูป?"
        .TextMatrix(0, 7) = "ประเภท Box"
        .TextMatrix(0, 8) = "จำนวน Box"
        .TextMatrix(0, 9) = "หมายเหตุ"
        .ColWidth(0) = 0
        .ColWidth(1) = 705
        .ColWidth(2) = 2550
        .ColWidth(3) = 1020
        .ColWidth(4) = 1020
        .ColWidth(5) = 1065
        .ColWidth(6) = 930
        .ColWidth(7) = 1000
        .ColWidth(8) = 950
        .ColWidth(9) = 2610
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
        '*** grCar ในรูปแบบ Grid Class
'        .ColComboList(5) = ComboBoxEnum
'        .ColComboList(6) = ComboBoxEnum
        
        '*** New VsFlexGrid
        .ColComboList(5) = ""
        .ColComboList(5) = TypeOfCar
        .ColComboList(6) = TypeOfPhotography
        .ColComboList(7) = TypeOfBox
        .Editable = flexEDKbd
        
    End With

    With grTimeTable.getGridObj
        .Rows = 2
        .Cols = 10
        .TextMatrix(0, 1) = "เที่ยวที่"
        .TextMatrix(0, 2) = "เวลา"
        .TextMatrix(0, 3) = "Vendor"
        .TextMatrix(0, 4) = "เลขทะเบียน"
        .TextMatrix(0, 5) = "คนขับ"
        .TextMatrix(0, 6) = "ประเภท Box"
        .TextMatrix(0, 7) = "จำนวน"
        .TextMatrix(0, 8) = "ลบบ็อก"
        .TextMatrix(0, 9) = "เลขใบคุม"
        .ColWidth(0) = 0
        .ColWidth(1) = 800
        .ColWidth(2) = 1185
        .ColWidth(3) = 2000
        .ColWidth(4) = 1560
        .ColWidth(5) = 1800
        .ColWidth(6) = 1000
        .ColWidth(7) = 700
        .ColWidth(8) = 800
        .ColWidth(9) = 500
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    grTimeTable.ColInputType(6) = ComboBoxEnum
'    grTimeTable.ColInputType(9) = ComboBoxEnum
    
End Sub

Private Sub CreateNewJobOpen()
    Set myJobOpenBox_ = Nothing
    Set myJobOpenBox_ = New BWGJobDataBox
    With myJobOpenBox_
        .ID = GetGUID
        Dim compID As String
            If cboCompany.ListIndex = 0 Then
                compID = "776EDA5830F744A6B3551348D851BC22"
            Else
                compID = "776EDA5830F744A6B3551348D851BC22"
            End If
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
        End If
        .IssuedDate = GetSystemDateTimeThai.SystemDate
        .IssuedTime = GetSystemDateTimeThai.SystemTime
        .IssuedStaffID = CurrentUser.ID
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .WorkDate = ctlWorkDate.ValueYMD
        .isChanged = False
    End With
End Sub

Private Sub grCar_AfterEdit(ByVal row As Long, ByVal col As Long)
Dim tmpDet As BWGJobDataCarUsedBox
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(row, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(row, 0))
    If isExist(myJobOpenBox_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsedBox
        tmpDet.ID = GetGUID
        tmpDet.JobID = myJobOpenBox_.ID

        grCar.TextMatrix(row, 0) = tmpDet.ID
        myJobOpenBox_.Cars.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myJobOpenBox_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(row, col))
    tmpDet.RowID = row
    With tmpDet
        Select Case col
            Case 3
                OldValue = Trim("" & .TruckCount)
                If IsNumeric(tmpVal) Then
                    .TruckCount = tmpVal
                Else
                    .TruckCount = 0
                End If
                If Trim(OldValue) <> Trim("" & .TruckCount) Then
                    .ClearTimeTable
                End If
            Case 4
                If IsNumeric(tmpVal) Then
                    .WorkerPerCar = tmpVal
                Else
                    .WorkerPerCar = 0
                End If
            Case 5
                OldValue = .GoSameTime
                If grCar.TextMatrix(row, col) = "1" Then
                    .GoSameTime = "Y"
                ElseIf grCar.TextMatrix(row, col) = "2" Then
                    .GoSameTime = "N"
                End If
                If Trim(OldValue) <> Trim(.GoSameTime) Then
                    If myJobOpenBox_.isCancel Or myJobOpenBox_.isCompleted = False Then
                        .ClearTimeTable
                        grTimeTable.ClearAllData
                    Else
                        MsgBox "ใบงานนี้ไม่สามารถแก้ไขได้"
                    End If
                End If
            Case 6
                If grCar.TextMatrix(row, col) = "1" Then
                    .TakePhoto = "W"
                ElseIf grCar.TextMatrix(row, col) = "2" Then
                    .TakePhoto = "B"
                ElseIf grCar.TextMatrix(row, col) = "3" Then
                    .TakePhoto = "F"
                ElseIf grCar.TextMatrix(row, col) = "4" Then
                    .TakePhoto = "N"
                End If
            Case 7
                .BoxType = grCar.TextMatrix(row, col)
            Case 8
                If Trim(grCar.TextMatrix(row, 7)) <> "" Then
                    BoxCheck = CInt(grCar.TextMatrix(row, 8))
                    If BoxCheck > BoxTypeCount(grCar.TextMatrix(row, 7)) Then
                        .BoxCount = 0
                        grCar.TextMatrix(row, 8) = ""
                        MsgBox "จำนวนที่ระบุมากกว่าจำนวนที่สามารถจัดบ็อกได้", vbInformation
                    Else
                        .BoxCount = BoxCheck
                    End If
                Else
                    grCar.TextMatrix(row, 8) = ""
                End If
                SelectBoxCount = tmpDet.BoxCount
            Case 9
                .Remark = Trim(tmpVal)
        End Select
        myJobOpenBox_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub grCar_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
Dim tmpCar As BWGJobDataCarUsedBox
    If (col = 3 Or col = 4) And grCar.TextMatrix(row, col) = "" Then Exit Sub
    If (col = 3 Or col = 4) And IsNumeric(grCar.TextMatrix(row, col)) = False Then
        MsgBox "กรุณากรอกข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        grCar.TextMatrix(row, col) = ""
        Cancel = True
        Exit Sub
    ElseIf col = 3 Then
        Set tmpCar = myJobOpenBox_.Cars("'" & Trim(grCar.TextMatrix(row, 0)) & "'")
        If tmpCar.isTrailer And CLng(grCar.TextMatrix(row, col)) Mod 2 <> 0 Then
            MsgBox "ท่านเลือกรถพ่วงต้องกรอกข้อมูลจำนวนเที่ยวเป็นเลขคู่เท่านั้น", vbExclamation
            grCar.TextMatrix(row, col) = ""
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub grCar_Click()
    Select Case grCar.col
        Case 9
            grCar.Editable = flexEDKbdMouse
        Case Else
            grCar.Editable = flexEDKbdMouse
    End Select
End Sub

Private Sub grCar_DblClick()
Dim x%, carNo%, tmpKey$, tmpCarTypeID$
Dim isTrailer As Boolean, carKey$
Dim tmpCar As BWGJobDataCarUsedBox, hasThisCar As Boolean

    With grCar
'            If Trim(.TextMatrix(.RowSel - 1, 2)) <> "" And .Col = 2 And Trim(.TextMatrix(.RowSel, 0)) <> "" Then
'                carNo = .TextMatrix(.RowSel, 1)
'                carKey = Trim(.TextMatrix(.RowSel, 0))
'                If isExist(myJobOpenBox_.JobDetails, "'" & carKey & "'") Then
'                    tmpCarTypeID = Trim(myJobOpenBox_.JobDetails("'" & carKey & "'").QuoTruckTypeID)
'                    If Trim(myJobOpenBox_.Cars("'" & carKey & "'").TruckSubTypeID) = "" Then
'                        isTrailer = Trim(myJobOpenBox_.JobDetails("'" & tmpKey & "'").isTrailer)
'                    Else
'                        isTrailer = Trim(myJobOpenBox_.Cars("'" & carKey & "'").isTrailer)
'                    End If
'                Else
'                    tmpCarTypeID = ""
'                End If
'                Load frmCarTypeSearch
'                With frmCarTypeSearch
'                    .SetCallerForm Me, tmpCarTypeID, , isTrailer
'                    .Show 1
'                End With
'            End If
            
            If .ColSel = 2 Then
                If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
                If Trim(grCar.TextMatrix(.RowSel, .ColSel)) = "" And Trim(curCarType_.TypeDesc) <> Trim(grCar.TextMatrix(.RowSel, .ColSel)) Then
                    Set curCarType_ = Nothing
                    Set curCarType_ = New BWGTruckType
                Else
                    If Trim(grCar.TextMatrix(.RowSel, .ColSel)) = "" Or Trim(curCarType_.TypeDesc) <> Trim(grCar.TextMatrix(.RowSel, .ColSel)) Then
                        Set curCarType_ = Nothing
                        Set curCarType_ = New BWGTruckType
                        If Not curCarType_ Is Nothing Then
'                            grCar.TextMatrix(Row, 0) = tmpDet.ID
                            grCar.TextMatrix(.RowSel, 1) = .RowSel
                            grCar.TextMatrix(.RowSel, .ColSel) = curCarType_.TypeDesc

                            If grCar.TextMatrix(grCar.row, 1) > 0 Then
                                hasThisCar = False
                                For x = 1 To .Rows - 1
                                    If IsNumeric(.TextMatrix(x, 1)) Then
                                        If carNo = CInt(.TextMatrix(x, 1)) Then
                                            tmpKey = Trim(.TextMatrix(x, 0))
                                            hasThisCar = True
                                            Exit For
                                        End If
                                    End If
                                Next
                                If hasThisCar = False Then
                                    Set tmpCar = New BWGJobDataCarUsedBox
                                    tmpCar.ID = GetGUID
                                    tmpCar.carNo = Int(grCar.TextMatrix(.RowSel, 1))
                                    tmpCar.JobID = myJobOpenBox_.ID
'                                    tmpCar.isTrailer = .isTrailer
                                    myJobOpenBox_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
                                    grCar.TextMatrix(myJobOpenBox_.Cars.Count, 0) = tmpCar.ID
                                    grCar.TextMatrix(myJobOpenBox_.Cars.Count, 1) = tmpCar.carNo
                                End If
                            End If

                            carNo = Trim(.TextMatrix(.RowSel, 1))
                            carKey = Trim(.TextMatrix(.RowSel, 0))
                            With grCar
                                For x = 1 To .Rows - 1
                                    If IsNumeric(.TextMatrix(x, 1)) Then
                                        If carNo = CInt(.TextMatrix(x, 1)) Then
                                            tmpKey = Trim(.TextMatrix(x, 0))
                                            Exit For
                                        End If
                                    End If
                                Next
                            End With
                            
'                            If isExist(myJobOpenBox_.Cars, "'" & Trim(tmpKey) & "'") Then
'                                Set curCar_ = myJobOpenBox_.Cars("'" & Trim(tmpKey) & "'")
'                            Else
'                                Set curCar_ = Nothing
'                            End If
                            
                            Load frmCarTypeSearch
                            With frmCarTypeSearch
                                .SetCallerForm Me, tmpCarTypeID, , isTrailer
                                .Show 1
                            End With

                        Else
                            grCar.TextMatrix(.RowSel, 1) = ""
                            grCar.TextMatrix(.RowSel, .ColSel) = ""
                        End If
                    End If
                End If
            End If
    End With
End Sub

Private Sub grCar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrD
Dim x%, carInUsed As Boolean, isTrailer As Boolean
Dim carNo%, tmpKey As String, tmpCarTypeID As String, carKey As String
    With grCar
        If Trim(.TextMatrix(.row, 1)) = "" Or IsNumeric(.TextMatrix(.row, 1)) = False Then
            Exit Sub
        End If
        'If Trim(.TextMatrix(.row - 1, 2)) <> "" And Trim(.TextMatrix(.row, 2)) = "" And .col = 2 And Trim(.TextMatrix(.row, 0)) <> "" Then
        carNo = .TextMatrix(.row, 1)
        carKey = "" 'Trim(.TextMatrix(.Row, 0))
        If KeyCode = 46 Then
            If MsgBox("ท่านต้องการลบรถที่ใช้ขนกากของเสียคันนี้ออกจากการเปิดงานครั้งนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                tmpKey = "'" & Trim(.TextMatrix(.row, 0)) & "'"
                Set curCar_ = myJobOpenBox_.Cars(Trim(tmpKey))
                If curCar_.TimeTables.Count > 0 Then
                    For x = 1 To curCar_.TimeTables.Count
                        carKey = "'" & Trim(grTimeTable.getGridObj.TextMatrix(1, 0)) & "'"
                        curCar_.TimeTables.Remove carKey
                        grTimeTable.getGridObj.RemoveItem 1
                        grTimeTable.getGridObj.Rows = grTimeTable.getGridObj.Rows + 1
                    Next x
                End If
                myJobOpenBox_.Cars.Remove tmpKey
                Set curCar_ = Nothing
                .RemoveItem .row
            End If
        Else
'            If Trim(.TextMatrix(.Row - 1, 2)) <> "" And .Col = 2 And Trim(.TextMatrix(.Row, 0)) <> "" Then
'                If isExist(myJobOpenBox_.JobDetails, "'" & carKey & "'") Then
'                    tmpCarTypeID = Trim(myJobOpenBox_.JobDetails("'" & carKey & "'").QuoTruckTypeID)
'                    If Trim(myJobOpenBox_.Cars("'" & carKey & "'").TruckSubTypeID) = "" Then
'                        isTrailer = Trim(myJobOpenBox_.JobDetails("'" & carKey & "'").isTrailer)
'                    Else
'                        isTrailer = Trim(myJobOpenBox_.Cars("'" & carKey & "'").isTrailer)
'                    End If
'                Else
'                    tmpCarTypeID = ""
'                End If
'                Load frmCarTypeSearch
'                With frmCarTypeSearch
'                    .SetCallerForm Me, tmpCarTypeID, , isTrailer
'                    .Show 1
'                End With
'            End If
        End If
    End With
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub grCar_RowColChange()
Dim tmpKey As String
Dim tmpCar As BWGJobDataCarUsedBox
Dim x%

    grCar.Editable = flexEDNone
    Frame4.Caption = "ตารางเดินรถ"

    tmpKey = Trim(grCar.TextMatrix(grCar.RowSel, 0))
    grTimeTable.ClearAllData
    grTimeTable.getGridObj.Rows = 2

    If tmpKey = vbNullString Then Exit Sub
    If isExist(myJobOpenBox_.Cars, "'" & Trim(tmpKey) & "'") Then
        Set curCar_ = myJobOpenBox_.Cars("'" & Trim(tmpKey) & "'")
    Else
        Set curCar_ = Nothing
    End If

    If Trim(grCar.TextMatrix(grCar.RowSel, 0)) = "" Then
        grCar.Editable = flexEDNone
        Frame4.Caption = "ตารางเดินรถ"
    Else
        If isExist(myJobOpenBox_.Cars, "'" & tmpKey & "'") Then
            Set tmpCar = myJobOpenBox_.Cars("'" & tmpKey & "'")
            Frame4.Caption = "ตารางเดินรถ [คันที่ " & tmpCar.carNo & "]"
            DoEvents
            If tmpCar.TimeTables.Count = 0 Then
                grTimeTable.getGridObj.Rows = 2
            Else
                grTimeTable.getGridObj.Rows = tmpCar.TimeTables.Count
            End If
            PopulateTimeTalbe tmpCar
            If tmpCar.BoxType <> "" Then
                SelectBoxType = tmpCar.BoxType
                SelectBoxCount = tmpCar.BoxCount
                grTimeTable.getColObject(6).Clear
                grTimeTable.getColObject(6).AddItem ""
                grTimeTable.getColObject(6).AddItem SelectBoxType
            End If
            For x = 1 To tmpCar.TimeTables.Count
                With grTimeTable.getGridObj
                    grTimeTable.getGridObj.TextMatrix(x, 0) = tmpCar.TimeTables(x).JobDataCarID & "_" & tmpCar.TimeTables(x).TripNo
                    grTimeTable.getGridObj.TextMatrix(x, 1) = tmpCar.TimeTables(x).TripNo
                    grTimeTable.getGridObj.TextMatrix(x, 2) = tmpCar.TimeTables(x).StartTime
                    grTimeTable.getGridObj.TextMatrix(x, 3) = tmpCar.TimeTables(x).VendorName
                    grTimeTable.getGridObj.TextMatrix(x, 4) = tmpCar.TimeTables(x).CarRegisNo
                    grTimeTable.getGridObj.TextMatrix(x, 5) = tmpCar.TimeTables(x).DriverName
                    grTimeTable.getGridObj.TextMatrix(x, 6) = tmpCar.TimeTables(x).BoxType
                    grTimeTable.getGridObj.TextMatrix(x, 7) = tmpCar.TimeTables(x).BoxCount
                    grTimeTable.getGridObj.TextMatrix(x, 9) = tmpCar.TimeTables(x).TimeTableNo
                    Dim Boxs As Integer
                    Boxs = curCar_.TimeTables(x).BoxCount
                    If Boxs > 0 Then
                        .col = 8: .row = x
                        Image1.Stretch = True
                        Set .CellPicture = Image1
                        .CellPictureAlignment = flexAlignCenterCenter
                        DoEvents
                    End If
                    .col = 0
                End With
            Next
            Set tmpCar = Nothing
        End If
    End If
End Sub

Private Sub grTimeTable_Click(row As Integer, col As Integer)
    On Error GoTo ErrD
    If col <> 8 Then Exit Sub
    If grTimeTable.getGridObj.CellPicture = Empty Then Exit Sub
    If MsgBox("ยืนยันการลบการจอง Box" & vbCrLf & "หมายเหตุ : บ็อกที่ถูกจัดไปแล้วจะถูกยกเลิกไปด้วย", vbInformation + vbYesNo) = vbNo Then Exit Sub
    If Not curCar_ Is Nothing Then
        curCar_.TimeTables(row).BoxType = ""
        curCar_.TimeTables(row).BoxCount = 0
        curCar_.TimeTables(row).Save
        grTimeTable.getGridObj.TextMatrix(row, 6) = ""
        grTimeTable.getGridObj.TextMatrix(row, 7) = ""
        Set grTimeTable.getGridObj.CellPicture = Nothing
    Else
        MsgBox "ไม่ได้เลือกรถ", vbExclamation
        Exit Sub
    End If
    Exit Sub
ErrD:
    MsgBox err.Description & vbExclamation
    Exit Sub
End Sub

Private Sub grTimeTable_DblClick(row As Integer, col As Integer)
Dim x%
    If cmdSave.Enabled = True Then
        If col = 3 Then
            If curCar_ Is Nothing Or Trim(grTimeTable.getGridObj.TextMatrix(grTimeTable.getGridObj.row, 0)) = "" Then
                MsgBox "กรุณาเลือกรถคันที่ต้องการ", vbExclamation
                Exit Sub
            End If
            If Trim(curCar_.ID) = "" Then
                MsgBox "กรุณาเลือกรถคันที่ต้องการ", vbExclamation
                Exit Sub
            End If
            If curCar_.isTrailer Then
                Load frmTrailerCarSearchBox
                With frmTrailerCarSearchBox
                    .SetCar curCar_
                    .Show 1
                End With
                If curCar_.isChanged Then
                    For x = 1 To curCar_.TimeTables.Count
                        With grTimeTable.getGridObj
                            grTimeTable.getGridObj.TextMatrix(x, 0) = curCar_.TimeTables(x).JobDataCarID & "_" & curCar_.TimeTables(x).TripNo
                            grTimeTable.getGridObj.TextMatrix(x, 1) = curCar_.TimeTables(x).TripNo
                            grTimeTable.getGridObj.TextMatrix(x, 2) = curCar_.TimeTables(x).StartTime
                            grTimeTable.getGridObj.TextMatrix(x, 3) = curCar_.TimeTables(x).VendorName
                            grTimeTable.getGridObj.TextMatrix(x, 4) = curCar_.TimeTables(x).CarRegisNo
                            grTimeTable.getGridObj.TextMatrix(x, 5) = curCar_.TimeTables(x).DriverName
                            grTimeTable.getGridObj.TextMatrix(x, 6) = curCar_.TimeTables(x).BoxType
                            grTimeTable.getGridObj.TextMatrix(x, 7) = curCar_.TimeTables(x).BoxCount
                            Dim Boxs As Integer
                            Boxs = curCar_.TimeTables(x).BoxCount
                            If Boxs > 0 Then
                                .col = 8: .row = x
                                Image1.Stretch = True
                                Set .CellPicture = Image1
                                .CellPictureAlignment = flexAlignCenterCenter
                                DoEvents
                            End If
                            .col = 0
                        End With
                    Next
                End If
            Else
                Load frmCarVendorSearch
                With frmCarVendorSearch
                    .SetCallerForm Me, curCar_.TruckSubTypeID
                    .Show 1
                End With
            End If
            myJobOpenBox_.isChanged = True
        End If
    End If
End Sub

Private Sub grTimeTable_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x%, carInUsed As Boolean
Dim carNo%, tmpKey As String
Dim tmpDet As BWGJobDataTimeTableBox
Dim tSql As String
    If cmdSave.Enabled Then
        If KeyCode = 46 And Trim(grTimeTable.getGridObj.TextMatrix(grTimeTable.getGridObj.row, 0)) <> "" Then
            If MsgBox("ท่านต้องการลบข้อมูลการกำหนดรถสำหรับรถเที่ยวนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                With grTimeTable.getGridObj
                    tmpKey = "'" & .TextMatrix(grTimeTable.getGridObj.row, 0) & "'"
                    Set tmpDet = curCar_.TimeTables(tmpKey)
                    For x = 1 To curCar_.TimeTables.Count
                        If Trim(tmpDet.VendorTruckID) = Trim(curCar_.TimeTables(x).TrailTruckID) Then
                            curCar_.TimeTables(x).TrailTruckID = ""
                        End If
                    Next

                    tSql = "UPDATE tbJobDataTimeTableBox SET VendorTruckID = '', DriverName = '' , BoxType='', BoxCount=0" & vbCrLf
                    tSql = tSql & "WHERE     (VendorTruckID = '" & tmpDet.VendorTruckID & "') AND (JobDataCarID = '" & tmpDet.JobDataCarID & "')" & vbCrLf
                    tSql = tSql & "AND (TripNo = '" & tmpDet.TripNo & "')"
                    DBConnExc tSql

                    tmpDet.VendorID = ""
                    tmpDet.VendorName = ""
                    tmpDet.VendorTruckID = ""
                    tmpDet.CarRegisNo = ""
                    tmpDet.DriverName = ""
                    tmpDet.BoxType = ""
                    tmpDet.BoxCount = 0
                    myJobOpenBox_.isChanged = False
'                    .TextMatrix(.row, 0) = ""
'                    .TextMatrix(.row, 1) = ""
'                    .TextMatrix(.row, 2) = ""
                    .TextMatrix(.row, 3) = ""
                    .TextMatrix(.row, 4) = ""
                    .TextMatrix(.row, 5) = ""
                    .TextMatrix(.row, 6) = ""
                    .TextMatrix(.row, 7) = ""
                    .TextMatrix(.row, 8) = ""
                End With
            End If
        ElseIf grTimeTable.getGridObj.col = 3 Then
            If curCar_ Is Nothing Or Trim(grTimeTable.getGridObj.TextMatrix(grTimeTable.getGridObj.row, 0)) = "" Then
                MsgBox "กรุณาเลือกรถคันที่ต้องการ", vbExclamation
                Exit Sub
            End If
            If Trim(curCar_.ID) = "" Then
                MsgBox "กรุณาเลือกรถคันที่ต้องการ", vbExclamation
                Exit Sub
            End If
                If curCar_.isTrailer Then
                    Load frmTrailerCarSearchBox
                    With frmTrailerCarSearchBox
                        .SetCar curCar_
                        .Show 1
                    End With
                    If curCar_.isChanged Then
                        For x = 1 To curCar_.TimeTables.Count
                            With grTimeTable.getGridObj
                                grTimeTable.getGridObj.TextMatrix(x, 0) = curCar_.TimeTables(x).JobDataCarID & "_" & curCar_.TimeTables(x).TripNo
                                grTimeTable.getGridObj.TextMatrix(x, 1) = curCar_.TimeTables(x).TripNo
                                grTimeTable.getGridObj.TextMatrix(x, 2) = curCar_.TimeTables(x).StartTime
                                grTimeTable.getGridObj.TextMatrix(x, 3) = curCar_.TimeTables(x).VendorName
                                grTimeTable.getGridObj.TextMatrix(x, 4) = curCar_.TimeTables(x).CarRegisNo
                                grTimeTable.getGridObj.TextMatrix(x, 5) = curCar_.TimeTables(x).DriverName
                                grTimeTable.getGridObj.TextMatrix(x, 6) = curCar_.TimeTables(x).BoxType
                                grTimeTable.getGridObj.TextMatrix(x, 7) = curCar_.TimeTables(x).BoxCount
                                grTimeTable.getGridObj.TextMatrix(x, 8) = ""
                                Dim Boxs As Integer
                                Boxs = curCar_.TimeTables(x).BoxCount
                                If Boxs > 0 Then
                                    .col = 8: .row = x
                                    Image1.Stretch = True
                                    Set .CellPicture = Image1
                                    .CellPictureAlignment = flexAlignCenterCenter
                                    DoEvents
                                End If
                                .col = 0
                            End With
                        Next
                    End If
                Else
                    Load frmCarVendorSearch
                    With frmCarVendorSearch
                        .SetCallerForm Me, curCar_.TruckSubTypeID
                        .Show 1
                    End With
                End If
        End If
    End If
End Sub

Private Sub grTimeTable_RolColChanged(row As Integer, col As Integer)
    If cmdSave.Enabled = False Then
        grTimeTable.ColEnabled(col) = False
    Else
        Select Case col
            Case 1, 3, 4
                grTimeTable.ColEnabled(col) = False
            Case 2, 5, 6, 7
                grTimeTable.ColEnabled(col) = True
        End Select
    End If
End Sub

Private Sub grTimeTable_UpdateValue(row As Integer, col As Integer)
On Error Resume Next
Dim tmpDet As BWGJobDataTimeTableBox
Dim tmpKey As String
Dim tmpVal As String, x%, SetAllBox As Integer
    If curCar_ Is Nothing Then Exit Sub
    tmpVal = Trim(grTimeTable.getGridObj.TextMatrix(row, col))
    tmpKey = Trim(grTimeTable.getGridObj.TextMatrix(row, 0))
    Select Case col
        Case 2
            If curCar_.GoSameTime = "Y" Then
                With grTimeTable.getGridObj
                    For x = 1 To .Rows - 1
                        tmpKey = Trim(grTimeTable.getGridObj.TextMatrix(x, 0))
                        If isExist(curCar_.TimeTables, "'" & tmpKey & "'") = False Then
                            Set tmpDet = New BWGJobDataTimeTableBox
                            tmpDet.JobDataCarID = curCar_.ID
                            tmpDet.TripNo = grTimeTable.getGridObj.TextMatrix(x, 1)
                            grTimeTable.getGridObj.TextMatrix(x, 0) = tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo)
                            curCar_.TimeTables.Add tmpDet, "'" & tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo) & "'"
                        Else
                            Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
                        End If
                        tmpDet.StartTime = tmpVal
                        Set tmpDet = Nothing
                    Next
                End With
            Else
                tmpKey = Trim(grTimeTable.getGridObj.TextMatrix(row, 0))
                If isExist(curCar_.TimeTables, "'" & tmpKey & "'") = False Then
                    Set tmpDet = New BWGJobDataTimeTableBox
                    tmpDet.JobDataCarID = curCar_.ID
                    tmpDet.TripNo = grTimeTable.getGridObj.TextMatrix(row, 1)
                    grTimeTable.getGridObj.TextMatrix(row, 0) = tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo)
                    curCar_.TimeTables.Add tmpDet, "'" & tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo) & "'"
                Else
                    Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
                End If
                tmpDet.StartTime = tmpVal
            End If
        Case 3
        Case 4
        Case 5
            Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
            tmpDet.DriverName = Trim(tmpVal)
        Case 6
            Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
            tmpDet.BoxType = tmpVal
            If Trim(grTimeTable.getGridObj.TextMatrix(row, 7)) = "" Then
                tmpDet.BoxCount = 0
            Else
                tmpDet.BoxCount = CInt(grTimeTable.getGridObj.TextMatrix(row, 7))
            End If
        Case 7
            Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
            With grTimeTable.getGridObj
                If .TextMatrix(row, 6) <> "" Then
                    SetAllBox = 0
                    For x = 1 To .Rows - 1
                        If .TextMatrix(row, 6) <> "" And .TextMatrix(row, 6) = SelectBoxType Then
                            If IsNumeric(.TextMatrix(x, 7)) Then
                                SetAllBox = SetAllBox + CInt(.TextMatrix(x, 7))
                            End If
                        End If
                    Next x
                    If SetAllBox <= SelectBoxCount Then
                        If IsNumeric(.TextMatrix(row, 7)) Then
                            If CInt(.TextMatrix(row, 7)) < 0 Then
                                 .TextMatrix(row, 7) = ""
                            Else
                                tmpDet.BoxCount = .TextMatrix(row, 7)
                            End If
                        Else
                            tmpDet.BoxCount = 0
                            .TextMatrix(row, 7) = ""
                        End If
                    Else
                        MsgBox "กำหนดบ็อกมากกว่าจำนวนที่มีการจองออเดอร์ไว้ (บ็อกที่มีในออเดอร์ " & SelectBoxCount & ") โปรดตรวจสอบอีกครั้ง"
                        tmpDet.BoxCount = 0
                        .TextMatrix(row, 7) = ""
                    End If
                End If
            End With
    End Select
    myJobOpenBox_.isChanged = True
    Set tmpDet = Nothing
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myJobOpenBox_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myJobOpenBox_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myJobOpenBox_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim x%, tmpDet As BWGCustomerSite
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    cboSiteName.Clear

    With curCustomer_
        For x = 1 To .OtherRemarks.Count
            lvResult.ListItems.Add , "'" & .OtherRemarks(x).ID & "'", .OtherRemarks(x).Remdesc
        Next
        For x = 1 To .SiteAddresses.Count
            Set tmpDet = .SiteAddresses(x)
            cboSiteName.AddItem .SiteAddresses(x).Address & " " & .SiteAddresses(x).TumbolName & " " & .SiteAddresses(x).AumphurName & " " & .SiteAddresses(x).ProvinceName '& " " & .SiteAddresses(x).ZipCode
            Set tmpDet = Nothing
        Next
        If .SiteAddresses.Count = 1 Then cboSiteName.ListIndex = 0

        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myJobOpenBox_.CustomerID = .ID
            myJobOpenBox_.CustomerName = .CustomerName
            myJobOpenBox_.UnderStaffID = .UnderSaleStaffID
            myJobOpenBox_.SaleStaffName = .UnderSaleName
        End If
    End With
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myJobOpenBox_.Remdesc = Trim(txtRem.Text)
End Sub

Public Sub SetDataTruckType(selTruck As BWGTruckType, Optional isAssigned As Boolean = True)
    Set curCarType_ = selTruck
    With curCarType_
        TruckType = .TypeDesc
'        cboIsTrailer.Enabled = .CanBeTrailer
'        If cboIsTrailer.Enabled = False Then cboIsTrailer.ListIndex = 0
        If isAssigned Then
'            myQuotation_.TruckTypeID = .ID
'            myQuotation_.TruckTypeName = .TypeDesc
'            myQuotation_.isTrailer = (cboIsTrailer.ListIndex = 1)
        End If
    End With
End Sub

Private Sub SetEnabledScreen()
Dim aa As Boolean
    aa = myJobOpenBox_.isCancel
    cmdSave.Enabled = Not aa And Not myJobOpenBox_.isCompleted
    cmdCancel.Enabled = Not aa And Not myJobOpenBox_.isCompleted
    cmdCloseJob.Enabled = Not aa And Not myJobOpenBox_.isCompleted
    cmdMenifest.Enabled = Not aa And Not myJobOpenBox_.isCompleted
    lbCancel.Visible = aa Or myJobOpenBox_.isCompleted
    If aa Then
        lbCancel.Caption = "ยกเลิก"
    ElseIf myJobOpenBox_.isCompleted Then
        lbCancel.Caption = "ปิด Job"
    End If
End Sub

Private Sub ClearScreen()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdMenifest.Enabled = True
    txtJobOpenNo.Text = "== AUTO =="
    ctlWorkDate.ValueDMY = FormatYMD_to_DMY(TodayDate, "/", "/")
    txtCustomer.Text = ""
    cboSiteName.Clear
    lvResult.ListItems.Clear
    Set curCustomer_ = Nothing
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Set curSaleStaff_ = Nothing
    txtRem.Text = ""
    grCar.Clear
    grCar.Rows = 20
    grTimeTable.ClearAllData
    grTimeTable.getGridObj.Rows = 2
    lbCancel.Visible = False
    Call SetGrid
    Set curCar_ = Nothing
End Sub

Public Sub SetDataJobOpen(selJob As BWGJobDataBox)
Dim tmpStr
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    SetGrid
    Set myJobOpenBox_ = selJob
    With myJobOpenBox_
        txtJobOpenNo.Text = .JobNo
        ctlWorkDate.ValueDMY = FormatYMD_to_DMY(.WorkDate, "/", "/")
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        SetTextToCombo .CustomerSiteName, cboSiteName
        If cboSiteName.ListCount = 1 Then
            If cboSiteName.ListIndex < 0 Then cboSiteName.ListIndex = 0
        End If
        If cboSiteName.ListCount > 0 Then
            If cboSiteName.ListIndex = -1 Then cboSiteName.ListIndex = 0
        End If
        txtIssuedStaff.Text = .IssuedStaffName
        SetTextToCombo CurrentUser.WorkCompanies("'" & .CompanyID & "'").CompanyNameTH, cboCompany
        cboCompany.Locked = True
        'Set tmpCol = tmpS.UserSearch("StaffID='" & .SaleStaffID & "'")
        txtRem.Text = .Remdesc
        
        If myJobOpenBox_.isNotActive = "Y" Then
            ChkNotActive.Value = 1
        Else
            ChkNotActive.Value = 0
        End If
        If myJobOpenBox_.JobType = "A" Then
            cboType.ListIndex = 0
        ElseIf myJobOpenBox_.JobType = "B" Then
            cboType.ListIndex = 1
        ElseIf myJobOpenBox_.JobType = "F" Then
            cboType.ListIndex = 2
        Else
            cboType.ListIndex = -1
        End If
'        cmdSave.Enabled = Not .MenifestCountAll = .MenifestCountIssued
        cmdCancel.Enabled = cmdSave.Enabled
        cmdMenifest.Enabled = cmdSave.Enabled
        SetEnabledScreen
        ShowCar
        ctlWorkDate.Enabled = False
        txtCustomer.Enabled = False
    End With
End Sub

Private Sub ShowCar()
Dim x%, tmpDet As BWGJobDataCarUsedBox
    With grCar
        grCar.Clear
        .Rows = myJobOpenBox_.Cars.Count + 20
        For x = 1 To myJobOpenBox_.Cars.Count
            Set tmpDet = myJobOpenBox_.Cars(x)
            .TextMatrix(x, 0) = tmpDet.ID
            .TextMatrix(x, 1) = tmpDet.carNo
            If tmpDet.isTrailer Then
                .TextMatrix(x, 2) = tmpDet.TruckTypeName & " (พ่วง), " & tmpDet.TruckSubTypeName
            Else
                .TextMatrix(x, 2) = tmpDet.TruckTypeName & ", " & tmpDet.TruckSubTypeName
            End If
            .TextMatrix(x, 3) = tmpDet.TruckCount
            .TextMatrix(x, 4) = tmpDet.WorkerPerCar
            If tmpDet.GoSameTime = "Y" Then
                .TextMatrix(x, 5) = "พร้อม"
            ElseIf tmpDet.GoSameTime = "N" Then
                .TextMatrix(x, 5) = "ไม่พร้อม"
            End If
            If tmpDet.TakePhoto = "Y" Then
                .TextMatrix(x, 6) = "ถ่าย"
            ElseIf tmpDet.TakePhoto = "N" Then
                .TextMatrix(x, 6) = "ไม่ถ่าย"
            ElseIf tmpDet.TakePhoto = "S" Then
                .TextMatrix(x, 6) = "ถ่ายรูปพิเศษ"
'            ElseIf tmpDet.TakePhoto = "W" Then
'                .TextMatrix(x, 6) = "ถ่ายรูปชั่ง + ลงของ"
'            ElseIf tmpDet.TakePhoto = "B" Then
'                .TextMatrix(x, 6) = "ถ่ายรูปขึ้นสายพาน"
'            ElseIf tmpDet.TakePhoto = "F" Then
'                .TextMatrix(x, 6) = "ถ่ายรูป Boxfeed"
            End If
            .TextMatrix(x, 7) = tmpDet.BoxType
            .TextMatrix(x, 8) = tmpDet.BoxCount
            .TextMatrix(x, 9) = tmpDet.Remark
            Set tmpDet = Nothing
        Next
    End With
End Sub

'Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
'Dim tmpCol As Collection
'Dim x%, isNew As Boolean
'Dim tmpID As String
'Dim tmpDetail As New BWGJobDataDetail
'    Set tmpCol = myJobOpenBox_.JobDetails
'
'    If tmpCol Is Nothing Then Set tmpCol = New Collection
'    isNew = True
'    tmpID = Trim(selItem.ID)
'    For x = 1 To tmpCol.Count
'        'If tmpCol(x).selitem.ID = tmpID Then
'        If tmpCol(x).wastedataID = tmpID Then
'            isNew = False
'            Exit For
'        End If
'    Next
'
'    If isNew Then
'        tmpDetail.ID = GetGUID
'        tmpDetail.JobID = myJobOpenBox_.ID
'        tmpDetail.wastedataID = selItem.ID
'        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
'        tmpDetail.WasteName = selItem.WasteName
'        tmpDetail.WasteNo = selItem.WasteNo
'        tmpDetail.WasteQuotationNo = selItem.QuotationNo
'        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
'        isDetailChange_ = True
'        myJobOpenBox_.isChanged = True
'        Set tmpDetail = Nothing
'    End If
'    If myJobOpenBox_.JobDetails Is Nothing Then
'        Set myJobOpenBox_.JobDetails = tmpCol
'        isDetailChange_ = True
'        myJobOpenBox_.isChanged = True
'    End If
'    Set tmpCol = Nothing
'End Sub
'
'Public Sub RemoveWasteData(selID As String)
'Dim tmpCol As Collection
'Dim x%
'Dim tmpKey As String
'    Set tmpCol = myJobOpenBox_.JobDetails
'    If tmpCol Is Nothing Then Exit Sub
'    For x = 1 To tmpCol.Count
'        If tmpCol(x).wastedataID = selID Then
'            tmpKey = "'" & tmpCol(x).ID & "'"
'            Exit For
'        End If
'    Next
'    If isExist(tmpCol, tmpKey) Then
'        tmpCol.Remove tmpKey
'        isDetailChange_ = True
'        myJobOpenBox_.isChanged = True
'    End If
'End Sub

'Public Function getWastDataItems() As Collection
'Dim x%
'Dim tmpCol As New Collection
'Dim tmpDet As BWGJobDataDetail
'Dim tmpWaste As BWGWasteDataCR
'Dim selCol As Collection, tmpStr As String, tmpArr
'    Set selCol = myJobOpenBox_.JobDetails
'    With selCol
'        For x = 1 To .Count
'            Set tmpDet = .Item(x)
'            If Trim(tmpDet.wastedataID) <> "" Then
'                Set tmpWaste = New BWGWasteDataCR
'                tmpWaste.ID = tmpDet.wastedataID
'                tmpStr = tmpDet.WasteEUCode
'                If InStr(1, tmpStr, "-") > 0 Then
'                    tmpArr = Split(tmpStr, "-")
'                    tmpWaste.EUCode1 = tmpArr(0)
'                    tmpWaste.EUCode2 = tmpArr(1)
'                    tmpWaste.EUCode3 = tmpArr(2)
'                End If
'                tmpWaste.WasteName = tmpDet.WasteName
'                tmpWaste.WasteNo = tmpDet.WasteNo
'                tmpWaste.QuotationNo = tmpDet.WasteQuotationNo
'                tmpCol.Add tmpWaste, "'" & tmpDet.wastedataID & "'"
'                Set tmpWaste = Nothing
'            End If
'            Set tmpDet = Nothing
'        Next
'    End With
'
'    Set getWastDataItems = tmpCol
'    Set tmpCol = Nothing
'End Function

Public Sub SetDataTruckSubType(selSubType As BWGTruckSubType, isTrailer As Boolean)
'Dim tmpCar As New BWGJobDataCarUsedBox
Dim x%, oldTypeID As String
Dim tmpTimeTable As BWGJobDataTimeTableBox
'    With grCar.getGridObj
'        For x = 1 To .Rows - 1
'            If Trim(.TextMatrix(x, 0)) = "" Then Exit For
'        Next
'    End With
    If curCar_ Is Nothing Then
        If isExist(myJobOpenBox_.Cars, "'" & grCar.TextMatrix(grCar.row, 0) & "'") Then
            Set curCar_ = myJobOpenBox_.Cars("'" & Trim(grCar.TextMatrix(grCar.row, 0)) & "'")
        End If
    End If
    With curCar_
        '.ID = GetGUID
        '.JobID = myJobOpenBox_.ID
        '.RowID = x
        If Trim(.TruckSubTypeID) <> Trim(selSubType.ID) Then
            myJobOpenBox_.isChanged = True
            For x = 1 To .TimeTables.Count
                Set tmpTimeTable = .TimeTables(x)
                tmpTimeTable.CarRegisNo = ""
                tmpTimeTable.DriverName = ""
                tmpTimeTable.VendorID = ""
                tmpTimeTable.VendorName = ""
                tmpTimeTable.VendorRegisNo = ""
                tmpTimeTable.VendorTruckID = ""
                tmpTimeTable.VendorWorkerCharge = 0
                Set tmpTimeTable = Nothing
                grTimeTable.getGridObj.TextMatrix(x, 3) = ""
                grTimeTable.getGridObj.TextMatrix(x, 4) = ""
                grTimeTable.getGridObj.TextMatrix(x, 5) = ""
            Next
        End If
        .TruckSubTypeID = selSubType.ID
        .TruckSubTypeName = selSubType.SubTypeDesc
        .TruckTypeName = selSubType.TypeDesc
        .isTrailer = isTrailer
        myJobOpenBox_.isChanged = True
    End With
    'myJobOpenBox_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
    With grCar
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 0)) = Trim(curCar_.ID) Then
                If curCar_.isTrailer Then
                    .TextMatrix(x, 2) = curCar_.TruckTypeName & " (พ่วง), " & curCar_.TruckSubTypeName
                Else
                    .TextMatrix(x, 2) = curCar_.TruckTypeName & ", " & curCar_.TruckSubTypeName
                End If
                Exit For
            End If
        Next
    End With
End Sub

Private Sub PopulateTimeTalbe(selCar As BWGJobDataCarUsedBox)
Dim x%
    With grTimeTable.getGridObj
'        If selCar.GoSameTime = "Y" Then
'            .Rows = 2
'            .TextMatrix(1, 1) = 1
'        ElseIf selCar.GoSameTime = "N" Then
            .Rows = selCar.TruckCount + 1
            For x = 1 To .Rows - 1
                .TextMatrix(x, 1) = x
            Next
'        End If
    End With
End Sub

Public Sub SetDataCarVendor(selCar As BWGVenderCar)
Dim tmpDet As BWGJobDataTimeTableBox
Dim tmpKey As String
    tmpKey = "'" & grTimeTable.getGridObj.TextMatrix(grTimeTable.getGridObj.row, 0) & "'"
    Set tmpDet = curCar_.TimeTables(tmpKey)
    tmpDet.VendorID = selCar.VendorID
    tmpDet.VendorName = selCar.VendorName
    tmpDet.VendorRegisNo = selCar.VendorLicenseNo
    tmpDet.VendorTruckID = selCar.ID
    tmpDet.VendorWorkerCharge = selCar.VendorWorkerCharge
    tmpDet.CarRegisNo = selCar.CarRegisID
    tmpDet.DriverName = selCar.ContactName
    myJobOpenBox_.isChanged = True
    With grTimeTable.getGridObj
        .TextMatrix(.row, 3) = tmpDet.VendorName
        .TextMatrix(.row, 4) = tmpDet.CarRegisNo
        .TextMatrix(.row, 5) = tmpDet.DriverName
    End With
    Set tmpDet = Nothing
End Sub

Private Sub LoadBoxType()
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer
    tSql = "SELECT     TOP (100) PERCENT BoxType, BoxTypeDesc, SUM(BoxStock) AS BoxStock" & vbCrLf
    tSql = tSql & "From dbo.vw_BoxQty" & vbCrLf
    tSql = tSql & "WHERE (BoxStatus = 1)" & vbCrLf
    tSql = tSql & "GROUP BY BoxType, BoxTypeDesc"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        If i = 1 Then
            TypeOfBox = "#" & rs!BoxType & "*" & i & ";" & rs!BoxTypeDesc & vbTab & rs!BoxType & vbTab & rs!BoxStock
        Else
            TypeOfBox = TypeOfBox & "|#" & rs!BoxType & ";" & rs!BoxTypeDesc & vbTab & rs!BoxType & vbTab & rs!BoxStock
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Function BoxTypeCount(ByVal BoxType As String) As Integer
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer
    Dim Reserve As Integer, BoxCount As Integer
    BoxTypeCount = 0
    tSql = "SELECT SUM(BoxStock) AS BoxStock" & vbCrLf
    tSql = tSql & "FROM dbo.vw_BoxQty" & vbCrLf
    tSql = tSql & "WHERE (BoxStatus = 1) AND BoxType = '" & BoxType & "'" & vbCrLf
    tSql = tSql & "GROUP BY BoxType"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    If Not rs.EOF Then
        BoxCount = rs!BoxStock
    Else
        BoxCount = 0
    End If
    Set rs = Nothing

    tSql = "SELECT SUM(BoxReservNumber) AS Reserve" & vbCrLf
    tSql = tSql & "From BWG_BOX.dbo.Box_BoxReserve" & vbCrLf
    tSql = tSql & "WHERE BoxType = '" & BoxType & "'" & vbCrLf
    tSql = tSql & "GROUP BY BoxType"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    If Not rs.EOF Then
        Reserve = rs!Reserve
    Else
        Reserve = 0
    End If

    BoxTypeCount = BoxCount - Reserve
End Function
