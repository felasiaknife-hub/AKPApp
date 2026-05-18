VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmJobBoxAssign 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F039-Box"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12435
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdComplete 
      Caption         =   "ปิด Job"
      Enabled         =   0   'False
      Height          =   795
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8340
      Width           =   975
   End
   Begin VB.CommandButton cmdManageBox 
      Caption         =   "จัด Box"
      Height          =   795
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7500
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "ตารางการเดินรถ"
      Height          =   2115
      Left            =   60
      TabIndex        =   31
      Top             =   7080
      Width           =   10275
      Begin WasteManagment.ctlGrid grTimeTable 
         Height          =   1725
         Left            =   60
         TabIndex        =   10
         Top             =   270
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3043
      End
      Begin VB.Shape Shape2 
         Height          =   1785
         Left            =   30
         Top             =   240
         Width           =   10155
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "รถที่ใช้ในการขน"
      Height          =   1635
      Left            =   60
      TabIndex        =   29
      Top             =   5460
      Width           =   12255
      Begin WasteManagment.ctlGrid grCar 
         Height          =   1335
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   2355
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3525
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   1830
      Top             =   30
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
      Height          =   2835
      Left            =   60
      TabIndex        =   20
      Top             =   780
      Width           =   12255
      Begin VB.TextBox txtWasteType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   690
         Width           =   1305
      End
      Begin VB.TextBox txtCustSite 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   690
         Width           =   4125
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   5145
      End
      Begin VB.TextBox txtIssuedDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1305
      End
      Begin VB.TextBox txtIssuedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   2595
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   4125
      End
      Begin VB.TextBox txtJobOpenNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   300
         Width           =   1755
      End
      Begin MSComctlLib.ListView lvResult 
         Height          =   1275
         Left            =   60
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1500
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         Height          =   375
         Left            =   9060
         Top             =   1050
         Width           =   1935
      End
      Begin VB.Label lbMenifestCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Left            =   9150
         TabIndex        =   33
         Top             =   1140
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menifest ที่ออกได้/ที่ออกแล้ว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   6990
         TabIndex        =   32
         Top             =   1140
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   28
         Top             =   1140
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
         Left            =   5820
         TabIndex        =   27
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ออกใบสั่งงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   24
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   6420
         TabIndex        =   23
         Top             =   360
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
         TabIndex        =   22
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
         Left            =   3870
         TabIndex        =   21
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9510
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8970
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8340
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   10470
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Enabled         =   0   'False
      Height          =   795
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7500
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "รายการกากของเสีย"
      Height          =   1875
      Left            =   60
      TabIndex        =   18
      Top             =   3600
      Width           =   12255
      Begin WasteManagment.ctlGrid grWasteData 
         Height          =   1575
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   2778
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   11550
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   945
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
      TabIndex        =   26
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "รายการรออนุมัติ, หรืออนุมัติแล้วเมื่อวันที่..... โดย...."
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   7830
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   3555
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
      Caption         =   "กำหนด Box ให้กับใบงาน"
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
      Width           =   2115
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
      Picture         =   "frJobBoxAssign.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmJobBoxAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myJobOpen_ As BWGJobData
Dim curCustomer_ As BWGCustomer
Dim curSaleStaff_ As BWGCompanyStaff
Dim isDetailChange_ As Boolean
Dim curCar_ As BWGJobDataCarUsed

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex > 0 Then
        myJobOpen_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
    Else
        myJobOpen_.companyID = ""
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim result
    result = MsgBox("ต้องการยกเลิกใบสั่งซื้อใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myJobOpen_.CancelJob
        If myJobOpen_.isCancel Then
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
        SetEnabledScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myJobOpen_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myJobOpen_.isChanged And cmdSave.Enabled Then Exit Sub
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

'***** ปิดการใช้งาน Procedure ออกไปเมนิเฟส
'Private Sub cmdMenifest_Click()
'Dim selResult As Integer
'Dim x%, Y%, isCompleted As Boolean
'Dim itemCount%, tmpcar As BWGJobDataCarUsed, tmpTimeTable As BWGJobDataTimeTable
'    If Trim(txtCustomer.Text) = "" Then
'        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
'        txtCustomer.SetFocus
'        Exit Sub
'    End If
'    If myJobOpen_.JobDetails Is Nothing Then
'        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
'        Exit Sub
'    End If
'
''    For x = 1 To myJobOpen_.Cars.Count
''        Set tmpCar = myJobOpen_.Cars(x)
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
'    isCompleted = myJobOpen_.isCompleted
'    If myJobOpen_.isChanged And cmdSave.Enabled Then
'        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
'        If selResult = vbYes Then
'            cmdSave_Click
'            If myJobOpen_.isChanged And cmdSave.Enabled Then Exit Sub
'        Else
'            MsgBox "ยังไม่ได้ทำการออกใบ Manifest", vbInformation
'            Exit Sub
'        End If
'    End If
'    Screen.MousePointer = 11
'    itemCount = myJobOpen_.GenerateMenifest
'    Screen.MousePointer = 0
'    MsgBox "เสร็จสิ้นการสร้างใบ Manifest จำนวน " & itemCount
'    With myJobOpen_
'        lbMenifestCount.Caption = .MenifestCountAll & "/" & .MenifestCountIssued
'        If isCompleted = False And .MenifestCountAll = .MenifestCountIssued Then
'            myJobOpen_.CloseJob
'        End If
'        cmdSave.Enabled = Not .MenifestCountAll = .MenifestCountIssued
'        cmdCancel.Enabled = cmdSave.Enabled
'        cmdMenifest.Enabled = cmdSave.Enabled
'        SetEnabledScreen
'    End With
'End Sub

Private Sub cmdManageBox_Click()
    Dim tmpS As New BWGSearchManager
    Dim tmpCol As Collection, tmpCar As BWGJobDataCarUsed
    Dim JobID As String, tmpKey As String
    If grCar.getGridObj.TextMatrix(grCar.getGridObj.RowSel, 0) = "" Then MsgBox "กรุณาเลือกรถที่ใช้ในการขน": Exit Sub
    JobID = myJobOpen_.ID
    Set tmpCol = tmpS.JobDataHeaderSearch("JobID='" & JobID & "'")
    If tmpCol Is Nothing Then Exit Sub
    
    tmpKey = Trim(grCar.getGridObj.TextMatrix(grCar.getGridObj.RowSel, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") Then
        Set tmpCar = myJobOpen_.Cars("'" & tmpKey & "'")
    Else
        MsgBox "กรุณาเลือกรถที่ใช้ในการขน"
        Exit Sub
    End If
    If tmpCol.Count > 0 Then
        Load frmBoxTransferOut
        With frmBoxTransferOut
            .SetDataBoxOpen tmpCol(1), tmpCar
            .Show 1
        End With
    End If
End Sub

Private Sub cmdSave_Click()
Dim x%, Y%, isCarInUse As Boolean
Dim tmpCol As Collection
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If myJobOpen_.JobDetails Is Nothing Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpen_.JobDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpen_.Cars Is Nothing Then
        MsgBox "กรุณาระบุรถที่ต้องใช้ในการขนอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpen_.Cars.Count = 0 Then
        MsgBox "กรุณาระบุรถที่ต้องใช้ในการขนอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    For x = 1 To myJobOpen_.Cars.Count
        If myJobOpen_.Cars(x).GoSameTime = "" Then
            MsgBox "กรุณาระบุว่ารถไปพร้อมกันหรือไม่ในแต่ละเที่ยวให้ครบถ้วน", vbExclamation
            Exit Sub
        End If
        For Y = 1 To myJobOpen_.Cars(x).TimeTables.Count
            If Trim(myJobOpen_.Cars(x).TimeTables(Y).StartTime) = "" Then
                MsgBox "กรุณาระบุตารางการเดินรถของรถคันที่ " & myJobOpen_.Cars(x).carNo & " ให้ครบถ้วน", vbExclamation
                Exit Sub
            End If
        Next
    Next
    With grWasteData.getGridObj
        For x = 1 To myJobOpen_.Cars.Count
            isCarInUse = False
            For Y = 1 To .Rows - 1
                If IsNumeric(Trim(.TextMatrix(Y, 8))) Then
                    If myJobOpen_.Cars(x).carNo = CInt(Trim(.TextMatrix(Y, 8))) Then
                        isCarInUse = True
                        Exit For
                    End If
                End If
            Next Y
            If isCarInUse = False Then
                MsgBox "รถคันที่ " & myJobOpen_.Cars(x).carNo & " ไม่ได้ถูกใช้งานแล้ว กรุณาลบออกจากรายการรถที่ใช้ขน", vbExclamation
                grCar.SetFocus
                Exit Sub
            End If
        Next x
    End With
    
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And myJobOpen_.isCompleted = False Then
        Set tmpCol = myJobOpen_.Cars
        For x = 1 To myJobOpen_.Cars.Count
            Set tmpCol = myJobOpen_.Cars(x).TimeTables
            DBConnExc "Update tbJobDataCarUsed Set TruckSubTypeID='" & myJobOpen_.Cars(x).TruckSubTypeID & "',isTrailer='" & TF_To_YN(myJobOpen_.Cars(x).isTrailer) & "' Where JobDataCarID='" & myJobOpen_.Cars(x).ID & "'"
            For Y = 1 To tmpCol.Count
                If Trim(tmpCol(Y).VendorTruckID) = "" Then
                    MsgBox "ข้อมูลผิดพลาด กรุณาจดเลขที่ Job แล้วติดต่อผู้ดูแลระบบทันที", vbExclamation
                    Exit Sub
                Else
                    DBConnExc "Update tbJobDataTimeTable Set VendorTruckID='" & tmpCol(Y).VendorTruckID & "',DriverName='" & tmpCol(Y).DriverName & "',TrailTruckID='" & tmpCol(Y).TrailTruckID & "' Where JobDataCarID='" & tmpCol(Y).JobDataCarID & "' And TripNo=" & tmpCol(Y).tripNo
                End If
            Next
            For Y = 1 To tmpCol.Count
                If Trim(tmpCol(Y).TrailTruckID) <> "" Then
                    DBConnExc "Update tbJobDataTimeTable Set TimeTableNo='" & tmpCol(Y).TimeTableNo & "' Where JobDataCarID='" & tmpCol(Y).JobDataCarID & "' And TripNo=" & (tmpCol(Y).tripNo + 1) & " And VendorTruckID='" & tmpCol(Y).TrailTruckID & "'"
                End If
            Next
        Next

        myJobOpen_.isChanged = False
    End If
    If myJobOpen_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGJobDataDetail
Dim tmpCol As Collection
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And myJobOpen_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myJobOpen_.isChanged And cmdSave.Enabled Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmJobOpenSearch
    With frmJobOpenSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdComplete_Click()
Dim result
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And cmdSave.Enabled Then
        result = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If result = vbYes Then
            cmdSave_Click
            If myJobOpen_.isChanged Then Exit Sub
        Else
            MsgBox "ยังไม่ได้ทำการปิด Job", vbExclamation
            Exit Sub
        End If
    End If
    result = MsgBox("ต้องการปิด Job นี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myJobOpen_.CloseJob
        SetEnabledScreen
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
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If

    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    PopulateUnit
    PopulateCompany
    txtIssuedDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    
    With grWasteData.getGridObj
        .Rows = 20
        .Cols = 10
        grWasteData.ColInputType(4) = ComboBoxEnum
        grWasteData.ColInputType(8) = ComboBoxEnum
        .TextMatrix(0, 1) = "Waste No."
        .TextMatrix(0, 2) = "ชื่อของเสีย"
        .TextMatrix(0, 3) = "ปริมาณของเสีย"
        .TextMatrix(0, 4) = "หน่วย"
        .TextMatrix(0, 5) = "บรรจุภัณฑ์"
        .TextMatrix(0, 6) = "จำนวนบรรจุภัณฑ์"
        .TextMatrix(0, 7) = "ปลายทางบ่อกำจัด"
        .TextMatrix(0, 8) = "รถคันที่"
        .TextMatrix(0, 9) = "หมายเหตุ"
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 3300
        .ColWidth(3) = 1185
        .ColWidth(4) = 1065
        .ColWidth(5) = 1230
        .ColWidth(6) = 1320
        .ColWidth(7) = 1545
        .ColWidth(8) = 1000
        .ColWidth(9) = 1245
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    grCar.ColInputType(5) = ComboBoxEnum
    grCar.ColInputType(6) = ComboBoxEnum
    With grCar.getGridObj
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
    End With
    With grTimeTable.getGridObj
        .Rows = 2
        .Cols = 6
        .TextMatrix(0, 1) = "เที่ยวที่"
        .TextMatrix(0, 2) = "เวลา"
        .TextMatrix(0, 3) = "Vendor"
        .TextMatrix(0, 4) = "เลขทะเบียน"
        .TextMatrix(0, 5) = "คนขับ"
        .ColWidth(0) = 0
        .ColWidth(1) = 1050
        .ColWidth(2) = 1185
        .ColWidth(3) = 3135
        .ColWidth(4) = 1560
        .ColWidth(5) = 2805
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    
    CreateNewJobOpen
End Sub

Private Sub CreateNewJobOpen()
    Set myJobOpen_ = Nothing
    Set myJobOpen_ = New BWGJobData
    With myJobOpen_
        .ID = GetGUID
        If cboCompany.ListIndex > 0 Then .companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        .IssuedDate = TodayDate
        .IssuedStaffID = CurrentUser.ID
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .WorkDate = TodayDate
        .isChanged = False
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myJobOpen_ Is Nothing Then Exit Sub
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And cmdSave.Enabled And myJobOpen_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myJobOpen_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim MinWidth As Integer
Dim NewWidth As Integer, NewHeight As Integer
    If Me.Width < 12555 Then Me.Width = 12555
    If Me.Height < 9705 Then Me.Height = 9705
    
    NewWidth = (Me.Width - 12555)
    NewHeight = Me.Height - 9705
    
    cmdManageBox.Left = Me.Width - cmdManageBox.Width - 150
    'cmdMenifest.Left = Me.Width - cmdMenifest.Width - 150
    cmdClose.Left = Me.Width - cmdClose.Width - 150
    cmdSave.Left = Me.Width - (cmdSave.Width * 2) - 200
    cmdComplete.Left = Me.Width - (cmdComplete.Width * 2) - 200
    
    cmdClose.Top = Me.Height - cmdClose.Height - 550
    cmdComplete.Top = Me.Height - cmdComplete.Height - 550
    cmdManageBox.Top = Me.Height - (cmdManageBox.Height * 2) - 600
    'cmdMenifest.Top = Me.Height - (cmdMenifest.Height * 2) - 600
    cmdSave.Top = Me.Height - (cmdSave.Height * 2) - 600

    MinWidth = 200
    Frame1.Width = Me.Width - MinWidth
    Frame2.Width = Me.Width - MinWidth
    Frame3.Width = Me.Width - MinWidth
    Frame4.Width = Me.Width - (cmdManageBox.Width + cmdSave.Width + 350)
    'Frame4.Width = Me.Width - (cmdMenifest.Width + cmdSave.Width + 350)
    Frame4.Height = 2115 + NewHeight
    
    Shape2.Width = Frame4.Width - 50
    Shape2.Height = 1785 + NewHeight
    
    lvResult.Width = Frame1.Width - 200
    grWasteData.Width = Frame2.Width - 100
    grCar.Width = Frame3.Width - 100
    grTimeTable.Width = Frame4.Width - 100
    grTimeTable.Height = 1725 + NewHeight
    
    lvResult.ColumnHeaders(1).Width = lvResult.Width
    grWasteData.getGridObj.ColWidth(9) = 1245 + NewWidth
    grCar.getGridObj.ColWidth(9) = 2610 + NewWidth
    grTimeTable.getGridObj.ColWidth(5) = 2805 + NewWidth
End Sub

Private Sub grCar_BeforeUpdateValue(newValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 3 Or col = 4) And IsNumeric(newValue) = False Then
        MsgBox "กรุณากรอกข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grCar_DblClick(row As Integer, col As Integer)
Dim tmpKey As String
Dim tmpCar As BWGJobDataCarUsed
    If cmdSave.Enabled And Trim(grCar.getGridObj.TextMatrix(row, 0)) <> "" Then
        With grCar.getGridObj
            tmpKey = "'" & .TextMatrix(row, 0) & "'"
            Set tmpCar = myJobOpen_.Cars(tmpKey)
            If Trim(.TextMatrix(.row - 1, 2)) <> "" And .col = 2 And Trim(.TextMatrix(.row, 0)) <> "" Then
                Load frmCarTypeSearch
                With frmCarTypeSearch
                    .SetCallerForm Me, tmpCar.TruckTypeID, , tmpCar.isTrailer
                    .Show 1
                End With
            End If
            Set tmpCar = Nothing
        End With
    End If
End Sub

Private Sub grCar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x%, carInUsed As Boolean
Dim carNo%, tmpKey As String
    If cmdSave.Enabled Then
        With grCar.getGridObj
            'If Trim(.TextMatrix(.row - 1, 2)) <> "" And Trim(.TextMatrix(.row, 2)) = "" And .col = 2 And Trim(.TextMatrix(.row, 0)) <> "" Then
            If KeyCode = 46 Then
                carNo = .TextMatrix(.row, 1)
                With grWasteData.getGridObj
                    For x = 1 To .Rows - 1
                        If IsNumeric(.TextMatrix(x, 8)) Then
                            If carNo = CInt(.TextMatrix(x, 8)) Then
                                MsgBox "ไม่สามารถลบรถคันนี้ออกจากการเปิดงานได้เพราะมีการเรียกใช้งานอยู่", vbExclamation
                                Exit Sub
                            End If
                        End If
                    Next
                End With
                If MsgBox("ท่านต้องการลบรถที่ใช้ขนกากของเสียคันนี้ออกจากการเปิดงานครั้งนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                    tmpKey = "'" & Trim(.TextMatrix(.row, 0)) & "'"
                    myJobOpen_.Cars.Remove tmpKey
                    .RemoveItem .row
                    grCar_RolColChanged .row, .col
                End If
            Else
                If Trim(.TextMatrix(.row - 1, 2)) <> "" And .col = 2 And Trim(.TextMatrix(.row, 0)) <> "" Then
                    tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                    Load frmCarTypeSearch
                    With frmCarTypeSearch
                        .SetCallerForm Me, myJobOpen_.Cars(tmpKey).TruckTypeID, , myJobOpen_.Cars(tmpKey).isTrailer
                        .Show 1
                    End With
                End If
            End If
        End With
    End If
End Sub

Private Sub grCar_LostFocus()
    grCar_RolColChanged grCar.getGridObj.row, grCar.getGridObj.col
End Sub

Private Sub grCar_RolColChanged(row As Integer, col As Integer)
Dim tmpKey As String
Dim tmpCar As BWGJobDataCarUsed
Dim x%
    
    grCar.ColEnabled(col) = False
    Frame4.Caption = "ตารางเดินรถ"

    tmpKey = Trim(grCar.getGridObj.TextMatrix(row, 0))
    grTimeTable.ClearAllData
    grTimeTable.getGridObj.Rows = 2
            
    If isExist(myJobOpen_.Cars, "'" & Trim(tmpKey) & "'") Then
        Set curCar_ = myJobOpen_.Cars("'" & Trim(tmpKey) & "'")
    Else
        Set curCar_ = Nothing
    End If
    
    If Trim(grCar.getGridObj.TextMatrix(row, 0)) = "" Then
        grCar.ColEnabled(col) = False
        Frame4.Caption = "ตารางเดินรถ"
    Else
        If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") Then
            Set tmpCar = myJobOpen_.Cars("'" & tmpKey & "'")
            Frame4.Caption = "ตารางเดินรถ [คันที่ " & tmpCar.carNo & "]"
            DoEvents
            If tmpCar.TimeTables.Count = 0 Then
                grTimeTable.getGridObj.Rows = 2
            Else
                grTimeTable.getGridObj.Rows = tmpCar.TimeTables.Count
            End If
            PopulateTimeTalbe tmpCar
            For x = 1 To tmpCar.TimeTables.Count
                grTimeTable.getGridObj.TextMatrix(x, 0) = tmpCar.TimeTables(x).JobDataCarID & "_" & tmpCar.TimeTables(x).tripNo
                grTimeTable.getGridObj.TextMatrix(x, 1) = tmpCar.TimeTables(x).tripNo
                grTimeTable.getGridObj.TextMatrix(x, 2) = tmpCar.TimeTables(x).StartTime
                grTimeTable.getGridObj.TextMatrix(x, 3) = tmpCar.TimeTables(x).VendorName
                grTimeTable.getGridObj.TextMatrix(x, 4) = tmpCar.TimeTables(x).CarRegisNo
                grTimeTable.getGridObj.TextMatrix(x, 5) = tmpCar.TimeTables(x).DriverName
            Next
            Set tmpCar = Nothing
        End If
        
    End If

End Sub

Private Sub grCar_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.getGridObj.TextMatrix(row, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.ID = GetGUID
        tmpDet.JobID = myJobOpen_.ID
        grCar.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.getGridObj.TextMatrix(row, col))
    tmpDet.rowID = row
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
                If grCar.getColObject(col).ListIndex = 0 Then
                    .GoSameTime = "Y"
                ElseIf grCar.getColObject(col).ListIndex = 1 Then
                    .GoSameTime = "N"
                End If
                If Trim(OldValue) <> Trim(.GoSameTime) Then
                    .ClearTimeTable
                End If
            Case 6
                If grCar.getColObject(col).ListIndex = 0 Then
                    .TakePhoto = "Y"
                ElseIf grCar.getColObject(col).ListIndex = 1 Then
                    .TakePhoto = "N"
                End If
            Case 7
                .BoxType = grCar.getGridObj.TextMatrix(row, col)
            Case 8
                If Trim(grCar.getGridObj.TextMatrix(row, 7)) <> "" Then
                    .BoxCount = CInt(grCar.getGridObj.TextMatrix(row, col))
                End If
            Case 9
                .Remark = Trim(tmpVal)
        End Select
        myJobOpen_.isChanged = True
    End With
    Set tmpDet = Nothing
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
                Load frmTrailerCarSearch
                With frmTrailerCarSearch
                    .SetCar curCar_
                    .Show 1
                End With
                If curCar_.isChanged Then
                    For x = 1 To curCar_.TimeTables.Count
                        grTimeTable.getGridObj.TextMatrix(x, 0) = curCar_.TimeTables(x).JobDataCarID & "_" & curCar_.TimeTables(x).tripNo
                        grTimeTable.getGridObj.TextMatrix(x, 1) = curCar_.TimeTables(x).tripNo
                        grTimeTable.getGridObj.TextMatrix(x, 2) = curCar_.TimeTables(x).StartTime
                        grTimeTable.getGridObj.TextMatrix(x, 3) = curCar_.TimeTables(x).VendorName
                        grTimeTable.getGridObj.TextMatrix(x, 4) = curCar_.TimeTables(x).CarRegisNo
                        grTimeTable.getGridObj.TextMatrix(x, 5) = curCar_.TimeTables(x).DriverName
                    Next
                End If
            Else
                Load frmCarVendorSearch
                With frmCarVendorSearch
                    .SetCallerForm Me, curCar_.TruckSubTypeID
                    .Show 1
                End With
            End If
            myJobOpen_.isChanged = True
        End If
    End If
End Sub

Private Sub grTimeTable_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x%, carInUsed As Boolean
Dim carNo%, tmpKey As String
Dim tmpDet As BWGJobDataTimeTable
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
                    
                    tSql = "UPDATE tbJobDataTimeTable SET VendorTruckID = '', DriverName = ''" & vbCrLf
                    tSql = tSql & "WHERE     (VendorTruckID = '" & tmpDet.VendorTruckID & "') AND (JobDataCarID = '" & tmpDet.JobDataCarID & "')" & vbCrLf
                    tSql = tSql & "AND (TripNo = '" & tmpDet.tripNo & "')"
                    DBConnExc tSql
                    
                    tmpDet.VendorID = ""
                    tmpDet.VendorName = ""
                    tmpDet.VendorTruckID = ""
                    tmpDet.CarRegisNo = ""
                    tmpDet.DriverName = ""
                    myJobOpen_.isChanged = False
'                    .TextMatrix(.row, 0) = ""
'                    .TextMatrix(.row, 1) = ""
'                    .TextMatrix(.row, 2) = ""
                    .TextMatrix(.row, 3) = ""
                    .TextMatrix(.row, 4) = ""
                    .TextMatrix(.row, 5) = ""
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
                    Load frmTrailerCarSearch
                    With frmTrailerCarSearch
                        .SetCar curCar_
                        .Show 1
                    End With
                    If curCar_.isChanged Then
                        For x = 1 To curCar_.TimeTables.Count
                            grTimeTable.getGridObj.TextMatrix(x, 0) = curCar_.TimeTables(x).JobDataCarID & "_" & curCar_.TimeTables(x).tripNo
                            grTimeTable.getGridObj.TextMatrix(x, 1) = curCar_.TimeTables(x).tripNo
                            grTimeTable.getGridObj.TextMatrix(x, 2) = curCar_.TimeTables(x).StartTime
                            grTimeTable.getGridObj.TextMatrix(x, 3) = curCar_.TimeTables(x).VendorName
                            grTimeTable.getGridObj.TextMatrix(x, 4) = curCar_.TimeTables(x).CarRegisNo
                            grTimeTable.getGridObj.TextMatrix(x, 5) = curCar_.TimeTables(x).DriverName
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
        If col = 1 Or col = 2 Or col = 3 Or col = 4 Then
            grTimeTable.ColEnabled(col) = False
        ElseIf col = 5 Then
            grTimeTable.ColEnabled(col) = True
        End If
    End If
End Sub

Private Sub grTimeTable_UpdateValue(row As Integer, col As Integer)
On Error Resume Next
Dim tmpDet As BWGJobDataTimeTable
Dim tmpKey As String
Dim tmpVal As String, x%
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
                            Set tmpDet = New BWGJobDataTimeTable
                            tmpDet.JobDataCarID = curCar_.ID
                            tmpDet.tripNo = grTimeTable.getGridObj.TextMatrix(x, 1)
                            grTimeTable.getGridObj.TextMatrix(x, 0) = tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.tripNo)
                            curCar_.TimeTables.Add tmpDet, "'" & tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.tripNo) & "'"
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
                    Set tmpDet = New BWGJobDataTimeTable
                    tmpDet.JobDataCarID = curCar_.ID
                    tmpDet.tripNo = grTimeTable.getGridObj.TextMatrix(row, 1)
                    grTimeTable.getGridObj.TextMatrix(row, 0) = tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.tripNo)
                    curCar_.TimeTables.Add tmpDet, "'" & tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.tripNo) & "'"
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
    End Select
    myJobOpen_.isChanged = True
    Set tmpDet = Nothing
End Sub

Private Sub grWasteData_BeforeUpdateValue(newValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 3 Or col = 6) And IsNumeric(newValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grWasteData_RolColChanged(row As Integer, col As Integer)
    grWasteData.ColEnabled(col) = False
End Sub

Private Sub grWasteData_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGJobDataDetail
Dim tmpKey As String
Dim tmpVal As String
Dim x%, tmpCar As BWGJobDataCarUsed, hasThisCar As Boolean
    tmpKey = Trim(grWasteData.getGridObj.TextMatrix(row, 0))
    If isExist(myJobOpen_.JobDetails, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataDetail
        tmpDet.ID = GetGUID
        tmpDet.JobID = myJobOpen_.ID
        grWasteData.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        myJobOpen_.JobDetails.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myJobOpen_.JobDetails("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grWasteData.getGridObj.TextMatrix(row, col))
    tmpDet.rowID = row
    With tmpDet
        Select Case col
            Case 3
                If IsNumeric(tmpVal) Then
                    .EstWasteQty = tmpVal
                Else
                    .EstWasteQty = 0
                End If
            Case 4
                .WasteQtyUnitID = colUnit_(grWasteData.getColObject(4).ListIndex + 1).ID
                .WasteQtyUnitName = grWasteData.getColObject(4).Text
            Case 5
                .ContainerTypeDesc = Trim(tmpVal)
            Case 6
                If IsNumeric(tmpVal) Then
                    .ContainerCount = tmpVal
                Else
                    .ContainerCount = 0
                End If
            Case 7
                .DestinationDesc = Trim(tmpVal)
            Case 8
                If IsNumeric(tmpVal) Then
                    .carNo = tmpVal
                Else
                    .carNo = 0
                End If
                hasThisCar = False
                For x = 1 To myJobOpen_.Cars.Count
                    If .carNo = myJobOpen_.Cars(x).carNo Then
                        hasThisCar = True
                        Exit For
                    End If
                Next
                If hasThisCar = False Then
                    Set tmpCar = New BWGJobDataCarUsed
                    tmpCar.ID = GetGUID
                    tmpCar.carNo = .carNo
                    tmpCar.JobID = myJobOpen_.ID
                    myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
                    grCar.getGridObj.TextMatrix(myJobOpen_.Cars.Count, 0) = tmpCar.ID
                    grCar.getGridObj.TextMatrix(myJobOpen_.Cars.Count, 1) = tmpCar.carNo
                End If
            Case 9
                .RemDesc = Trim(tmpVal)
        End Select
        myJobOpen_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
'    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
'    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
'        Set curCustomer_ = Nothing
'        Set curCustomer_ = New BWGCustomer
'        With myJobOpen_
'            .CustomerID = ""
'            .CustomerName = ""
'        End With
'    Else
'        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
'            Set curCustomer_ = Nothing
'            Set curCustomer_ = New BWGCustomer
'            With myJobOpen_
'                .CustomerID = ""
'                .CustomerName = ""
'            End With
'            Load frmCustomerSearch
'            With frmCustomerSearch
'                .getCriteria (txtCustomer.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myJobOpen_.CustomerID = "" Then txtCustomer.Text = ""
'        End If
'    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim x%, tmpDet As BWGCustomerSite
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        For x = 1 To .OtherRemarks.Count
            lvResult.ListItems.Add , "'" & .OtherRemarks(x).ID & "'", .OtherRemarks(x).RemDesc
        Next
        For x = 1 To .SiteAddresses.Count
            Set tmpDet = .SiteAddresses(x)
            'cboCustSite.AddItem Trim(tmpDet.SiteName & " " & tmpDet.Address & " " & tmpDet.TumbolName & " " & tmpDet.AumphurName & " " & tmpDet.ProvinceName)
            Set tmpDet = Nothing
        Next
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myJobOpen_.CustomerID = .ID
            myJobOpen_.CustomerName = .CustomerName
        End If
    End With
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myJobOpen_.RemDesc = Trim(txtRem.Text)
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    cboCompany.AddItem "บริษัทที่ออกใบสั่งงานเก็บขน"
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount = 2 Then
        cboCompany.ListIndex = 1
    Else
        cboCompany.ListIndex = 0
    End If
End Sub

Private Sub SetEnabledScreen()
Dim JobCancel As Boolean
    JobCancel = myJobOpen_.isCancel
    cmdSave.Enabled = Not JobCancel And Not myJobOpen_.isCompleted
    cmdCancel.Enabled = Not JobCancel And Not myJobOpen_.isCompleted
    cmdComplete.Enabled = Not JobCancel And Not myJobOpen_.isCompleted
    cmdManageBox.Enabled = Not JobCancel And myJobOpen_.isCompleted
    'cmdMenifest.Enabled = Not aa And Not myJobOpen_.isCompleted
    lbCancel.Visible = JobCancel Or myJobOpen_.isCompleted
    If JobCancel Then
        lbCancel.Caption = "ยกเลิก"
    ElseIf myJobOpen_.isCompleted Then
        lbCancel.Caption = "ปิด Job"
    End If
End Sub

Private Sub ClearScreen()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdManageBox.Enabled = True
    'cmdSave.Enabled = True
    'cmdCancel.Enabled = True
    'cmdMenifest.Enabled = True
    txtJobOpenNo.Text = "== AUTO =="
    txtIssuedDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    txtCustomer.Text = ""
    lvResult.ListItems.Clear
    Set curCustomer_ = Nothing
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    txtWasteType.Text = ""
    Set curSaleStaff_ = Nothing
    txtRem.Text = ""
    grWasteData.ClearAllData
    grWasteData.getGridObj.Rows = 20
    grCar.ClearAllData
    grCar.getGridObj.Rows = 20
    grTimeTable.ClearAllData
    grTimeTable.getGridObj.Rows = 2
End Sub

Public Sub SetDataJobOpen(selJob As BWGJobData)
Dim tmpStr
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myJobOpen_ = selJob
    With myJobOpen_
        txtJobOpenNo.Text = .JobNo
        txtIssuedDate.Text = FormatYMD_to_DMY(.WorkDate, "/", "/")
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        txtCustSite.Text = .CustomerSiteName
        txtIssuedStaff.Text = .IssuedStaffName
        txtWasteType.Text = .WasteType
        SetTextToCombo CurrentUser.WorkCompanies("'" & .companyID & "'").CompanyNameTH, cboCompany
        cboCompany.Locked = True
        'Set tmpCol = tmpS.UserSearch("StaffID='" & .SaleStaffID & "'")
        txtRem.Text = .RemDesc
        lbMenifestCount.Caption = .MenifestCountAll & "/" & .MenifestCountIssued
        
        '*** ตรวจสอบเพิ่มเติมเกี่ยวกับ กรณีที่ยังไม่ได้ออกใบเมนิเฟส หรือออกยังไม่ครบ
        cmdSave.Enabled = Not .MenifestCountAll = .MenifestCountIssued
        cmdCancel.Enabled = cmdSave.Enabled
        cmdManageBox.Enabled = .MenifestCountAll = .MenifestCountIssued
        'cmdMenifest.Enabled = cmdSave.Enabled
        SetEnabledScreen
        ShowWasteData
        ShowCar
    End With
End Sub

Private Sub ShowWasteData()
Dim x%, tmpDet As BWGJobDataDetail
    With grWasteData.getGridObj
        grWasteData.ClearAllData
        .Rows = myJobOpen_.JobDetails.Count + 20
        For x = 1 To myJobOpen_.JobDetails.Count
            Set tmpDet = myJobOpen_.JobDetails(x)
            .TextMatrix(x, 0) = tmpDet.ID
            .TextMatrix(x, 1) = tmpDet.WasteNo
            .TextMatrix(x, 2) = tmpDet.WasteName & " [" & tmpDet.WasteType & "]"
            .TextMatrix(x, 3) = tmpDet.EstWasteQty
            .TextMatrix(x, 4) = tmpDet.WasteQtyUnitName
            .TextMatrix(x, 5) = tmpDet.ContainerTypeDesc
            .TextMatrix(x, 6) = tmpDet.ContainerCount
            .TextMatrix(x, 7) = tmpDet.TSDFName
            If tmpDet.carNo > 0 Then .TextMatrix(x, 8) = tmpDet.carNo
            .TextMatrix(x, 9) = tmpDet.RemDesc
            Set tmpDet = Nothing
        Next
    End With
End Sub

Private Sub ShowCar()
Dim x%, tmpDet As BWGJobDataCarUsed
    With grCar.getGridObj
        grCar.ClearAllData
        .Rows = myJobOpen_.Cars.Count + 20
        For x = 1 To myJobOpen_.Cars.Count
            Set tmpDet = myJobOpen_.Cars(x)
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
            End If
            .TextMatrix(x, 7) = tmpDet.BoxType
            .TextMatrix(x, 8) = tmpDet.BoxCount
            .TextMatrix(x, 9) = tmpDet.Remark
            Set tmpDet = Nothing
        Next
    End With
End Sub

Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGJobDataDetail
    Set tmpCol = myJobOpen_.JobDetails

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).wastedataID = tmpID Then
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpDetail.JobID = myJobOpen_.ID
        tmpDetail.wastedataID = selItem.ID
        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
        tmpDetail.WasteName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpDetail.WasteQuotationNo = selItem.QuotationNo
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
        isDetailChange_ = True
        myJobOpen_.isChanged = True
        Set tmpDetail = Nothing
    End If
    If myJobOpen_.JobDetails Is Nothing Then
        Set myJobOpen_.JobDetails = tmpCol
        isDetailChange_ = True
        myJobOpen_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveWasteData(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myJobOpen_.JobDetails
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).wastedataID = selID Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myJobOpen_.isChanged = True
    End If
End Sub

Public Function getWastDataItems() As Collection
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGJobDataDetail
Dim tmpWaste As BWGWasteDataCR
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myJobOpen_.JobDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.wastedataID) <> "" Then
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = tmpDet.wastedataID
                tmpStr = tmpDet.WasteEUCode
                If InStr(1, tmpStr, "-") > 0 Then
                    tmpArr = Split(tmpStr, "-")
                    tmpWaste.EUCode1 = tmpArr(0)
                    tmpWaste.EUCode2 = tmpArr(1)
                    tmpWaste.EUCode3 = tmpArr(2)
                End If
                tmpWaste.WasteName = tmpDet.WasteName
                tmpWaste.WasteNo = tmpDet.WasteNo
                tmpWaste.QuotationNo = tmpDet.WasteQuotationNo
                tmpCol.Add tmpWaste, "'" & tmpDet.wastedataID & "'"
                Set tmpWaste = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With

    Set getWastDataItems = tmpCol
    Set tmpCol = Nothing
End Function

Public Sub SetDataTruckSubType(selSubType As BWGTruckSubType, isTrailer As Boolean)
'Dim tmpCar As New BWGJobDataCarUsed
Dim x%, oldTypeID As String
Dim tmpTimeTable As BWGJobDataTimeTable
'    With grCar.getGridObj
'        For x = 1 To .Rows - 1
'            If Trim(.TextMatrix(x, 0)) = "" Then Exit For
'        Next
'    End With
    If curCar_ Is Nothing Then
        If isExist(myJobOpen_.Cars, "'" & Trim(grCar.getGridObj.TextMatrix(grCar.getGridObj.row, 0)) & "'") Then
            Set curCar_ = myJobOpen_.Cars("'" & Trim(grCar.getGridObj.TextMatrix(grCar.getGridObj.row, 0)) & "'")
        End If
    End If
    With curCar_
        '.ID = GetGUID
        '.JobID = myJobOpen_.ID
        '.RowID = x
        If Trim(.TruckSubTypeID) <> Trim(selSubType.ID) Then
            myJobOpen_.isChanged = True
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
        myJobOpen_.isChanged = True
    End With
    'myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
    With grCar.getGridObj
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

Private Sub PopulateTimeTalbe(selCar As BWGJobDataCarUsed)
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
Dim tmpDet As BWGJobDataTimeTable
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
    myJobOpen_.isChanged = True
    With grTimeTable.getGridObj
        .TextMatrix(.row, 3) = tmpDet.VendorName
        .TextMatrix(.row, 4) = tmpDet.CarRegisNo
        .TextMatrix(.row, 5) = tmpDet.DriverName
    End With
    Set tmpDet = Nothing
End Sub
