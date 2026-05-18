VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobCarAssign 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F039"
   ClientHeight    =   9480
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12660
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "รายการภาชนะ"
      Height          =   555
      Left            =   90
      TabIndex        =   41
      Top             =   3600
      Width           =   12210
      Begin VB.Label LblCon 
         AutoSize        =   -1  'True
         Caption         =   "ไม่มี"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   135
         TabIndex        =   42
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdMenifest 
      Caption         =   "ออก Manifest"
      Height          =   795
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ปิด Job"
      Height          =   795
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8520
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "ตารางการเดินรถ"
      Height          =   2100
      Left            =   60
      TabIndex        =   30
      Top             =   7695
      Width           =   10275
      Begin WasteManagment.ctlGrid grTimeTable 
         Height          =   1335
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   2355
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   7200
         Picture         =   "frJobCarAssign.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   315
      End
      Begin VB.Shape Shape2 
         Height          =   1455
         Left            =   60
         Top             =   240
         Width           =   10155
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
      TabIndex        =   19
      Top             =   720
      Width           =   12255
      Begin VB.CheckBox chkIsNot 
         BackColor       =   &H80000009&
         Caption         =   "งานขนส่งพิเศษ"
         Height          =   195
         Left            =   9420
         TabIndex        =   36
         Top             =   1140
         Width           =   1515
      End
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
         TabIndex        =   29
         Top             =   690
         Width           =   5295
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
         Width           =   4005
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
         Width           =   5295
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
         Width           =   7905
         _ExtentX        =   13944
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
      Begin VB.Label LblCarCheck 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9945
         TabIndex        =   39
         Top             =   1845
         Width           =   2220
      End
      Begin VB.Label LblEmpTrain 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9945
         TabIndex        =   38
         Top             =   1530
         Width           =   2220
      End
      Begin VB.Label lblDTeam2 
         BackColor       =   &H00FF00FF&
         Caption         =   "EAM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   10125
         TabIndex        =   35
         Top             =   495
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDTeam1 
         BackColor       =   &H00FF00FF&
         Caption         =   "DT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   435
         Left            =   10125
         TabIndex        =   34
         Top             =   75
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         Height          =   375
         Left            =   7920
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label lbMenifestCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Left            =   8040
         TabIndex        =   32
         Top             =   1140
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menifest ที่ออกได้/ที่ออกแล้ว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5820
         TabIndex        =   31
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8820
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8820
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "รายการกากของเสีย"
      Height          =   1740
      Left            =   60
      TabIndex        =   17
      Top             =   4230
      Width           =   12255
      Begin WasteManagment.ctlGrid grWasteData 
         Height          =   2070
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8820
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "รถที่ใช้ในการขน"
      Height          =   1635
      Left            =   60
      TabIndex        =   28
      Top             =   6015
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
   Begin VB.Label LblVIP 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "VIP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   6120
      TabIndex        =   40
      Top             =   45
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LbCustCode 
      Height          =   285
      Left            =   11250
      TabIndex        =   37
      Top             =   7380
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   10440
      Picture         =   "frJobCarAssign.frx":0442
      Stretch         =   -1  'True
      Top             =   7380
      Visible         =   0   'False
      Width           =   315
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
      Left            =   3870
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   18
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
      Picture         =   "frJobCarAssign.frx":0840
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmJobCarAssign"
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
Dim SelectBoxType As String
Dim SelectBoxCount As Integer

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myJobOpen_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).id
        Else
            myJobOpen_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
        End If
    Else
        myJobOpen_.CompanyID = ""
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
            CmdSave_Click
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

Private Sub cmdMenifest_Click()
On Error GoTo ErrD
Dim selResult As Integer
Dim x%, y%, isCompleted As Boolean
Dim itemCount%, tmpCar As BWGJobDataCarUsed, tmpTimeTable As BWGJobDataTimeTable
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If myJobOpen_.JobDetails Is Nothing Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    
'    For x = 1 To myJobOpen_.Cars.Count
'        Set tmpCar = myJobOpen_.Cars(x)
'        For y = 1 To tmpCar.TimeTables.Count
'            Set tmpTimeTable = tmpCar.TimeTables(y)
'            If Trim(tmpTimeTable.VendorTruckID) = "" Then
'                MsgBox "กรุณากำหนดรถให้ครบถ้วนก่อนทำการสร้างใบ Menifest", vbExclamation
'                Set tmpCar = Nothing
'                Set tmpTimeTable = Nothing
'                Exit Sub
'            End If
'            Set tmpTimeTable = Nothing
'        Next
'        Set tmpCar = Nothing
'    Next
    isCompleted = myJobOpen_.isCompleted
    If myJobOpen_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpen_.isChanged And cmdSave.Enabled Then Exit Sub
        Else
            MsgBox "ยังไม่ได้ทำการออกใบ Manifest", vbInformation
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    
    itemCount = myJobOpen_.GenerateMenifest
    Screen.MousePointer = 0
    MsgBox "เสร็จสิ้นการสร้างใบ Manifest จำนวน " & itemCount
    With myJobOpen_
        lbMenifestCount.Caption = .MenifestCountAll & "/" & .MenifestCountIssued
        If isCompleted = False And .MenifestCountAll = .MenifestCountIssued Then
            myJobOpen_.CloseJob
        End If
        cmdSave.Enabled = Not .MenifestCountAll = .MenifestCountIssued
        cmdCancel.Enabled = cmdSave.Enabled
        cmdMenifest.Enabled = cmdSave.Enabled
        SetEnabledScreen
    End With
    Exit Sub
ErrD:
    MsgBox "เกิดข้อผิดพลาดกรุณาลองใหม่", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub CmdSave_Click()
Dim x%, y%, isCarInUse As Boolean
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
        For y = 1 To myJobOpen_.Cars(x).TimeTables.Count
            If Trim(myJobOpen_.Cars(x).TimeTables(y).StartTime) = "" Then
                MsgBox "กรุณาระบุตารางการเดินรถของรถคันที่ " & myJobOpen_.Cars(x).carNo & " ให้ครบถ้วน", vbExclamation
                Exit Sub
            End If
        Next
    Next
    With grWasteData.getGridObj
        For x = 1 To myJobOpen_.Cars.Count
            isCarInUse = False
            For y = 1 To .Rows - 1
                If IsNumeric(Trim(.TextMatrix(y, 8))) Then
                    If myJobOpen_.Cars(x).carNo = CInt(Trim(.TextMatrix(y, 8))) Then
                        isCarInUse = True
                        Exit For
                    End If
                End If
            Next y
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
            DBConnExc "Update tbJobDataCarUsed Set TruckSubTypeID='" & myJobOpen_.Cars(x).TruckSubTypeID & "',isTrailer='" & TF_To_YN(myJobOpen_.Cars(x).isTrailer) & "' WHERE JobDataCarID='" & myJobOpen_.Cars(x).id & "'"
            For y = 1 To tmpCol.Count
                If Trim(tmpCol(y).VendorTruckID) = "" Then
                    MsgBox "ข้อมูลผิดพลาด กรุณาจดเลขที่ Job แล้วติดต่อผู้ดูแลระบบทันที", vbExclamation
                    Exit Sub
                Else
                    If Trim(tmpCol(y).BoxType) <> "" And tmpCol(y).BoxCount = 0 Then MsgBox "ไม่ได้กำหนดจำนวน box กรุณากำหนดจำนวน box", vbInformation
                    If Not tmpCol(y) Is Nothing Then
                        tmpCol(y).isChanged = True
                        tmpCol(y).Save
                    End If
'                    DBConnExc "Update tbJobDataTimeTable Set VendorTruckID='" & tmpCol(y).VendorTruckID & "',DriverName='" & tmpCol(y).DriverName & "', BoxType = '" & tmpCol(y).BoxType & "', BoxCount=" & tmpCol(y).BoxCount & ", LastUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "', LastUpdate = GetDate() WHERE JobDataCarID='" & tmpCol(y).JobDataCarID & "' And TripNo=" & tmpCol(y).TripNo
                End If
                If chkIsNot.Value = 1 Then
                                DBConnExc "Update tbJobDataHeader Set isNotActive = 'Y'  WHERE JobID='" & myJobOpen_.id & "' "
                                DBConnExc "Update tbJobDataTimeTable Set isNotActive = 'Y'  WHERE JobDataCarID='" & tmpCol(y).JobDataCarID & "' "
                End If
              
'                isNotActive
'                ''**** อัพเดทใบคุมกรณีรถพ่วงให้เป็นเลขเดียวกัน
'                If Trim(tmpCol(y).TrailTruckID) <> "" Then
'                    DBConnExc "Update tbJobDataTimeTable Set TimeTableNo='" & tmpCol(y).TimeTableNo & "' WHERE JobDataCarID='" & tmpCol(y).JobDataCarID & "' And TripNo=" & (tmpCol(y).TripNo + 1) & " And VendorTruckID='" & tmpCol(y).TrailTruckID & "'"
'                End If
                ''*******************************************
            Next y
        Next x
        myJobOpen_.isChanged = False
    End If
    DBConnExc "Update tbMenifestHeader Set UnderSaleStaffID = '" & curCustomer_.UnderSaleStaffID & "' , SaleDistrictID = '" & curCustomer_.SaleDistrictID & "'  WHERE CreatedFromJobID='" & myJobOpen_.id & "' "
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
            CmdSave_Click
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

Private Sub Command1_Click()
Dim result
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And cmdSave.Enabled Then
        result = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If result = vbYes Then
            CmdSave_Click
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
    PopulateCompany cboCompany, cboIndex
    txtIssuedDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    
    With grWasteData.getGridObj
        .Rows = 20
        .Cols = 10
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
    grWasteData.ColInputType(4) = ComboBoxEnum
    grWasteData.ColInputType(8) = ComboBoxEnum
    
    With grCar.getGridObj
         .Rows = 20
        .Cols = 17
        .TextMatrix(0, 1) = "รถคันที่"
        .TextMatrix(0, 2) = "ประเภทรถ"
        .TextMatrix(0, 3) = "จำนวนเที่ยว"
        .TextMatrix(0, 4) = "คนงาน/คัน"
        .TextMatrix(0, 5) = "ไปพร้อมกัน?"
        .TextMatrix(0, 6) = "ถ่ายรูป?"
        .TextMatrix(0, 7) = "Box"
        .TextMatrix(0, 8) = "จำนวน"
        .TextMatrix(0, 9) = "ปากกว้าง"
        .TextMatrix(0, 10) = "เอากลับ"
        .TextMatrix(0, 11) = "ปากจู๋"
        .TextMatrix(0, 12) = "เอากลับ"
        .TextMatrix(0, 13) = "ถุงบิกแบ็ค"
        .TextMatrix(0, 14) = "เบ้า"
        .TextMatrix(0, 15) = "ถุงดำ"
        .TextMatrix(0, 16) = "หมายเหตุ"
        .ColWidth(0) = 0
        .ColWidth(1) = 650
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1400
        .ColWidth(7) = 700
        .ColWidth(8) = 700
        .ColWidth(9) = 850
        .ColWidth(10) = 850
        .ColWidth(11) = 850
        .ColWidth(12) = 850
        .ColWidth(13) = 850
        .ColWidth(14) = 850
        .ColWidth(15) = 850
        .ColWidth(16) = 8000
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    grCar.ColInputType(5) = ComboBoxEnum
    grCar.ColInputType(6) = ComboBoxEnum
    With grTimeTable.getGridObj
        .Rows = 2
        .Cols = 11
        .TextMatrix(0, 1) = "เที่ยวที่"
        .TextMatrix(0, 2) = "เวลา"
        .TextMatrix(0, 3) = "Vendor"
        .TextMatrix(0, 4) = "เลขทะเบียน"
        .TextMatrix(0, 5) = "คนขับ"
        .TextMatrix(0, 6) = "ประเภท Box"
        .TextMatrix(0, 7) = "จำนวน"
        .TextMatrix(0, 8) = "ลบบ็อก"
        .TextMatrix(0, 9) = "..."
        .TextMatrix(0, 10) = "Rate น้ำมัน"
        .ColWidth(0) = 0
        .ColWidth(1) = 800
        .ColWidth(2) = 1185
        .ColWidth(3) = 2000
        .ColWidth(4) = 1560
        .ColWidth(5) = 1800
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 800
        .ColWidth(9) = 500
        .ColWidth(10) = 800
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    grTimeTable.ColInputType(6) = ComboBoxEnum
    grTimeTable.ColInputType(9) = ComboBoxEnum
    CreateNewJobOpen
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim MinWidth As Integer
Dim NewWidth As Integer, NewHeight As Integer
    If Me.Width < 12555 Then Me.Width = 12555
    If Me.Height < 9705 Then Me.Height = 9705
    
    NewWidth = (Me.Width - 12555)
    NewHeight = Me.Height - 9705
    
    cmdMenifest.Left = Me.Width - cmdMenifest.Width - 150
    cmdClose.Left = Me.Width - cmdClose.Width - 150
    cmdSave.Left = Me.Width - (cmdSave.Width * 2) - 200
    Command1.Left = Me.Width - (Command1.Width * 2) - 200
    
    cmdClose.Top = Me.Height - cmdClose.Height - 550
    Command1.Top = Me.Height - Command1.Height - 550
    cmdMenifest.Top = Me.Height - (cmdMenifest.Height * 2) - 600
    cmdSave.Top = Me.Height - (cmdSave.Height * 2) - 600

    MinWidth = 200
    Frame1.Width = Me.Width - MinWidth
    Frame2.Width = Me.Width - MinWidth
    Frame3.Width = Me.Width - MinWidth
    Frame4.Width = Me.Width - (cmdMenifest.Width + cmdSave.Width + 350)
    Frame5.Width = Me.Width - MinWidth
    Frame4.Height = 1915 + NewHeight
    
    Shape2.Width = Frame4.Width - 50
    Shape2.Height = 1785 + NewHeight
    LblEmpTrain.Left = Me.Width - 2600
    LblCarCheck.Left = Me.Width - 2600
    lvResult.Width = Frame1.Width - 2700
    grWasteData.Width = Frame2.Width - 100
    grCar.Width = Frame3.Width - 100
    grTimeTable.Width = Frame4.Width - 100
    grTimeTable.Height = 1125 + NewHeight
    
    lvResult.ColumnHeaders(1).Width = lvResult.Width
    grWasteData.getGridObj.ColWidth(9) = 1245 + NewWidth
'    grCar.getGridObj.ColWidth(9) = 2610 + NewWidth
    grTimeTable.getGridObj.ColWidth(5) = 1800 + NewWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myJobOpen_ Is Nothing Then Exit Sub
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And cmdSave.Enabled And myJobOpen_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpen_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub CreateNewJobOpen()
    Set myJobOpen_ = Nothing
    Set myJobOpen_ = New BWGJobData
    With myJobOpen_
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
        .IssuedDate = TodayDate
        .IssuedStaffID = CurrentUser.id
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .WorkDate = TodayDate
        .isChanged = False
    End With
End Sub

Private Sub grCar_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 3 Or col = 4) And IsNumeric(NewValue) = False Then
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
On Error GoTo nextStep
    grCar_RolColChanged grCar.getGridObj.row, grCar.getGridObj.col
nextStep:
End Sub

Private Sub grCar_RolColChanged(row As Integer, col As Integer)
'On Error GoTo nextStep

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
                    grTimeTable.getGridObj.TextMatrix(x, 8) = ""
                    grTimeTable.getGridObj.TextMatrix(x, 10) = tmpCar.TimeTables(x).OilUse
                    Dim Boxs As Integer
                    Boxs = curCar_.TimeTables(x).BoxCount
                    If Boxs > 0 Then
                        .col = 8: .row = x
                        Image3.Stretch = True
                        Image3.Height = 315
                        Image3.Height = 315
                        Set .CellPicture = Image3
                        .CellPictureAlignment = flexAlignCenterCenter
                        DoEvents
                    End If
                    .col = 0
                End With
            Next
            Set tmpCar = Nothing
        End If
    End If
'    End Sub
'nextStep:
'    End Sub
End Sub

Private Sub grCar_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.getGridObj.TextMatrix(row, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        grCar.getGridObj.TextMatrix(row, 0) = tmpDet.id
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.getGridObj.TextMatrix(row, col))
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
                    
                'ElseIf grCar.getColObject(col).ListIndex = 2 Then
                '    .TakePhoto = "W"
                    
                ElseIf grCar.getColObject(col).ListIndex = 2 Then
                    .TakePhoto = "S"
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
            If Trim(curCar_.id) = "" Then
                MsgBox "กรุณาเลือกรถคันที่ต้องการ", vbExclamation
                Exit Sub
            End If
'            If curCar_.isTrailer Then
'                Load frmTrailerCarSearch
'                With frmTrailerCarSearch
'                    .SetCar curCar_
'                    .Show 1
'                End With
'                If curCar_.isChanged Then
'                    For X = 1 To curCar_.TimeTables.Count
'                        With grTimeTable.getGridObj
'                            grTimeTable.getGridObj.TextMatrix(X, 0) = curCar_.TimeTables(X).JobDataCarID & "_" & curCar_.TimeTables(X).TripNo
'                            grTimeTable.getGridObj.TextMatrix(X, 1) = curCar_.TimeTables(X).TripNo
'                            grTimeTable.getGridObj.TextMatrix(X, 2) = curCar_.TimeTables(X).StartTime
'                            grTimeTable.getGridObj.TextMatrix(X, 3) = curCar_.TimeTables(X).VendorName
'                            grTimeTable.getGridObj.TextMatrix(X, 4) = curCar_.TimeTables(X).CarRegisNo
'                            grTimeTable.getGridObj.TextMatrix(X, 5) = curCar_.TimeTables(X).DriverName
'                            grTimeTable.getGridObj.TextMatrix(X, 6) = curCar_.TimeTables(X).BoxType
'                            grTimeTable.getGridObj.TextMatrix(X, 7) = curCar_.TimeTables(X).BoxCount
'                            Dim Boxs As Integer
'                            Boxs = curCar_.TimeTables(X).BoxCount
'                            If Boxs > 0 Then
'                                .col = 8: .row = X
'                                Image3.Stretch = True
'                                Image3.Height = 315
'                                Image3.Height = 315
'                                Set .CellPicture = Image3
'                                .CellPictureAlignment = flexAlignCenterCenter
'                                DoEvents
'                            End If
'                            .col = 0
'                        End With
'                    Next
'                End If
'            Else
                Load frmCarVendorSearch
                With frmCarVendorSearch
                    .SetCallerForm Me, curCar_.TruckSubTypeID
                    .Show 1
                End With
'            End If
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
                    
                    tSql = "UPDATE tbJobDataTimeTable SET VendorTruckID = '', DriverName = '' , BoxType='', BoxCount=0" & vbCrLf
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
                    myJobOpen_.isChanged = False
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
            If Trim(curCar_.id) = "" Then
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
                                Image3.Stretch = True
                                Image3.Height = 315
                                Image3.Height = 315
                                Set .CellPicture = Image3
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
        If col = 1 Or col = 2 Or col = 3 Or col = 4 Then
            grTimeTable.ColEnabled(col) = False
        ElseIf col = 5 Or col = 6 Or col = 7 Then
            grTimeTable.ColEnabled(col) = True
        End If
    End If
End Sub

Private Sub grTimeTable_UpdateValue(row As Integer, col As Integer)
On Error Resume Next
Dim tmpDet As BWGJobDataTimeTable
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
                            Set tmpDet = New BWGJobDataTimeTable
                            tmpDet.JobDataCarID = curCar_.id
                            tmpDet.TripNo = grTimeTable.getGridObj.TextMatrix(x, 1)
                            'tmpDet.TripTranFee = curCar_
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
                    Set tmpDet = New BWGJobDataTimeTable
                    tmpDet.JobDataCarID = curCar_.id
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
                            If IsNumeric(.TextMatrix(row, 7)) Then
                                SetAllBox = SetAllBox + .TextMatrix(row, 7)
                            End If
                        End If
                    Next x
'                    If SetAllBox <= SelectBoxCount Then
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
'                    Else
'                        MsgBox "กำหนดบ็อกมากกว่าจำนวนที่มีการจองออเดอร์ไว้ (บ็อกที่มีในออเดอร์ " & SelectBoxCount & ") โปรดตรวจสอบอีกครั้ง"
'                        tmpDet.BoxCount = 0
'                        .TextMatrix(Row, 7) = ""
'                    End If
                End If
            
                
            End With
            Case 10
                Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
                If Trim(grTimeTable.getGridObj.TextMatrix(row, 10)) = "" Then
                    tmpDet.OilUse = 0
                Else
                    tmpDet.OilUse = CInt(grTimeTable.getGridObj.TextMatrix(row, 10))
                End If
    End Select
    myJobOpen_.isChanged = True
    Set tmpDet = Nothing
End Sub

Private Sub grWasteData_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 3 Or col = 6) And IsNumeric(NewValue) = False Then
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
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        grWasteData.getGridObj.TextMatrix(row, 0) = tmpDet.id
        myJobOpen_.JobDetails.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.JobDetails("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grWasteData.getGridObj.TextMatrix(row, col))
    tmpDet.RowID = row
    With tmpDet
        Select Case col
            Case 3
                If IsNumeric(tmpVal) Then
                    .EstWasteQty = tmpVal
                Else
                    .EstWasteQty = 0
                End If
            Case 4
                .WasteQtyUnitID = colUnit_(grWasteData.getColObject(4).ListIndex + 1).id
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
                    tmpCar.id = GetGUID
                    tmpCar.carNo = .carNo
                    tmpCar.JobID = myJobOpen_.id
                    myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.id & "'"
                    grCar.getGridObj.TextMatrix(myJobOpen_.Cars.Count, 0) = tmpCar.id
                    grCar.getGridObj.TextMatrix(myJobOpen_.Cars.Count, 1) = tmpCar.carNo
                End If
            Case 9
                .Remdesc = Trim(tmpVal)
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
            lvResult.ListItems.Add , "'" & .OtherRemarks(x).id & "'", .OtherRemarks(x).Remdesc
        Next
        For x = 1 To .SiteAddresses.Count
            Set tmpDet = .SiteAddresses(x)
            'cboCustSite.AddItem Trim(tmpDet.SiteName & " " & tmpDet.Address & " " & tmpDet.TumbolName & " " & tmpDet.AumphurName & " " & tmpDet.ProvinceName)
            Set tmpDet = Nothing
        Next
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myJobOpen_.CustomerID = .id
            myJobOpen_.CustomerName = .CustomerName
        End If
        lblDTeam1.Visible = .DTeam
        lblDTeam2.Visible = .DTeam
        If .isVIPCR = "Y" Then
            LblVIP.Visible = True
        Else
            LblVIP.Visible = False
        End If
        LbCustCode.Caption = .CustomerCode
    End With
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myJobOpen_.Remdesc = Trim(txtRem.Text)
End Sub

Private Sub SetEnabledScreen()
Dim aa As Boolean
    aa = myJobOpen_.isCancel
    cmdSave.Enabled = Not aa And Not myJobOpen_.isCompleted
    cmdCancel.Enabled = Not aa And Not myJobOpen_.isCompleted
    Command1.Enabled = Not aa And Not myJobOpen_.isCompleted
    cmdMenifest.Enabled = Not aa And Not myJobOpen_.isCompleted
    lbCancel.Visible = aa Or myJobOpen_.isCompleted
    If aa Then
        lbCancel.Caption = "ยกเลิก"
    ElseIf myJobOpen_.isCompleted Then
        lbCancel.Caption = "ปิด Job"
    End If
End Sub

Private Sub ClearScreen()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdMenifest.Enabled = True
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
    lblDTeam1.Visible = False
    lblDTeam2.Visible = False
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
        SetTextToCombo CurrentUser.WorkCompanies("'" & .CompanyID & "'").CompanyNameTH, cboCompany
        cboCompany.Locked = True
        If myJobOpen_.isNotActive = "Y" Then
            chkIsNot.Value = 1
        Else
            chkIsNot.Value = 0
        End If
        'Set tmpCol = tmpS.UserSearch("StaffID='" & .SaleStaffID & "'")
        txtRem.Text = .Remdesc
        lbMenifestCount.Caption = .MenifestCountAll & "/" & .MenifestCountIssued
        
        cmdSave.Enabled = Not .MenifestCountAll = .MenifestCountIssued
        cmdCancel.Enabled = cmdSave.Enabled
        cmdMenifest.Enabled = cmdSave.Enabled
        SetEnabledScreen
        ShowWasteData
        ShowCar
        
        Dim i As Integer
            Dim rs As New ADODB.Recordset
            Dim x%, tSql As String
            Dim MS1 As String
                tSql = "SELECT * FROM tbprob WHERE custid='" & LbCustCode.Caption & "'  ORDER BY ProbDate"
                Set rs = New ADODB.Recordset
                Set rs = GetRS(tSql, adOpenStatic, adLockOptimistic)
                        If rs.RecordCount > 0 Then
                        MS1 = ""
                        For x = 1 To rs.RecordCount
                        MS1 = MS1 & rs!ProbName & " "
                                    rs.MoveNext
                        Next
                        MsgBox "ลูกค้านี้เคยมีข้อร้องเรียนเรื่อง " & MS1 & " กรุณาระมัดระวัง"
                    End If

    End With
End Sub

Private Sub ShowWasteData()
Dim x%, tmpDet As BWGJobDataDetail
    With grWasteData.getGridObj
        grWasteData.ClearAllData
        .Rows = myJobOpen_.JobDetails.Count + 20
        For x = 1 To myJobOpen_.JobDetails.Count
            Set tmpDet = myJobOpen_.JobDetails(x)
            .TextMatrix(x, 0) = tmpDet.id
            .TextMatrix(x, 1) = tmpDet.WasteNo
            .TextMatrix(x, 2) = tmpDet.WasteName & " [" & tmpDet.WasteType & "]"
            .TextMatrix(x, 3) = tmpDet.EstWasteQty
            .TextMatrix(x, 4) = tmpDet.WasteQtyUnitName
            .TextMatrix(x, 5) = tmpDet.ContainerTypeDesc
            .TextMatrix(x, 6) = tmpDet.ContainerCount
            .TextMatrix(x, 7) = tmpDet.TSDFName
            If tmpDet.carNo > 0 Then .TextMatrix(x, 8) = tmpDet.carNo
            .TextMatrix(x, 9) = tmpDet.Remdesc
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
            .TextMatrix(x, 0) = tmpDet.id
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
            'ElseIf tmpDet.TakePhoto = "W" Then
             '   .TextMatrix(x, 6) = "ถ่ายรูปชั่ง + ลงของ"
            ElseIf tmpDet.TakePhoto = "S" Then
                .TextMatrix(x, 6) = "ถ่ายรูปพิเศษ"
'            ElseIf tmpDet.TakePhoto = "B" Then
'                .TextMatrix(x, 6) = "ถ่ายรูปขึ้นสายพาน"
'            ElseIf tmpDet.TakePhoto = "F" Then
'                .TextMatrix(x, 6) = "ถ่ายรูป Boxfeed"
            End If
            .TextMatrix(x, 7) = tmpDet.BoxType
            .TextMatrix(x, 8) = tmpDet.BoxCount
             .TextMatrix(x, 9) = tmpDet.Tankcount
            .TextMatrix(x, 10) = tmpDet.TankTypeB
            .TextMatrix(x, 11) = tmpDet.TankTypeSCount
            .TextMatrix(x, 12) = tmpDet.TankTypeSB
            .TextMatrix(x, 13) = tmpDet.BagCount
            .TextMatrix(x, 14) = tmpDet.BagCount
            .TextMatrix(x, 15) = tmpDet.BagCount
            .TextMatrix(x, 16) = tmpDet.Remark
'            If tmpDet.TankType = "L" Then
'                .TextMatrix(X, 9) = "ปากกว้าง"
'            ElseIf tmpDet.TankType = "S" Then
'                .TextMatrix(X, 9) = "ปากแคบ"
'            End If
'            .TextMatrix(X, 10) = tmpDet.Tankcount
'            If tmpDet.BagType = "C" Then
'                .TextMatrix(X, 11) = "ปากป่อง"
'            ElseIf tmpDet.BagType = "P" Then
'                .TextMatrix(X, 11) = "เปิดปาก"
'            End If
'            .TextMatrix(X, 13) = tmpDet.BagCount
'            .TextMatrix(X, 14) = tmpDet.Remark
    Dim Cont As String
    If Val(tmpDet.Container) > 0 Then
        Cont = "เบ้าท์ใส่น้ำเสีย " & tmpDet.Container & " "
    End If
    If Val(tmpDet.BoutS) > 0 Then
        Cont = Cont & "เบ้าท์ตัดปาก " & tmpDet.BoutS & " "
    End If
    If Val(tmpDet.Lack) > 0 Then
        Cont = Cont & "แลครุ่นเก่า " & tmpDet.Lack & " "
    End If
    If Val(tmpDet.LackNew) > 0 Then
        Cont = Cont & "แลครุ่นใหม่ " & tmpDet.LackNew & " "
    End If
    If Val(tmpDet.LackGas) > 0 Then
        Cont = Cont & "แลคใส่ถังแแก๊ส " & tmpDet.LackGas & " "
    End If
    If Val(tmpDet.Plate) > 0 Then
        Cont = Cont & "พาเลท " & tmpDet.Plate & " "
    End If
    If Val(tmpDet.Plastic) > 0 Then
        Cont = Cont & "พลาสติกแล็ป " & tmpDet.Plastic & " "
    End If
    If Val(tmpDet.Tankcount) > 0 Then
        Cont = Cont & "ถัง200 ลิตรปากกว้าง " & tmpDet.Tankcount & " "
    End If
    If Val(tmpDet.TankTypeSCount) > 0 Then
        Cont = Cont & "ถัง200 ลิตรปากจู๋ " & tmpDet.TankTypeSCount & " "
    End If
    If Val(tmpDet.BagCount) > 0 Then
        Cont = Cont & "Bigbag " & tmpDet.BagCount & " "
    End If
    If Val(tmpDet.BagBack) > 0 Then
        Cont = Cont & "ถุงดำ " & tmpDet.BagBack & " "
    End If
    If Val(tmpDet.PaperBox) > 0 Then
        Cont = Cont & "กล่องกระดาษ " & tmpDet.PaperBox & " "
    End If
    If Val(tmpDet.BoutCage) > 0 Then
        Cont = Cont & "กรงเบ้าท์ " & tmpDet.BoutCage
    End If
    If Val(tmpDet.CovidBin) > 0 Then
        Cont = Cont & "ถังขยะโควิด 150   " & tmpDet.CovidBin
    End If
    If Val(tmpDet.CovidBin200) > 0 Then
        Cont = Cont & "ถังขยะโควิด  200    " & tmpDet.CovidBin200
    End If
    If Val(tmpDet.Rack) > 0 Then
        Cont = Cont & "ถังพลาสติก 200 ลิตรปากจู๋    " & tmpDet.Rack
    End If

    LblCon.Caption = Cont
         If tmpDet.EmpTrain = "Y" Then
            LblEmpTrain.Caption = "ต้องอบรมพนักงานก่อน"
        Else
            LblEmpTrain.Caption = ""
        End If
        If tmpDet.CarCheck = "Y" Then
            LblCarCheck.Caption = "ต้องตรวจสภาพรถก่อน"
        Else
            LblCarCheck.Caption = ""
        End If
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
        tmpDetail.JobID = myJobOpen_.id
        tmpDetail.wastedataID = selItem.id
        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
        tmpDetail.WasteName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpDetail.WasteQuotationNo = selItem.QuotationNo
        tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
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
            tmpKey = "'" & tmpCol(x).id & "'"
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
                tmpWaste.id = tmpDet.wastedataID
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
        If Trim(.TruckSubTypeID) <> Trim(selSubType.id) Then
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
        .TruckSubTypeID = selSubType.id
        .TruckSubTypeName = selSubType.SubTypeDesc
        .TruckTypeName = selSubType.TypeDesc
        .isTrailer = isTrailer
        myJobOpen_.isChanged = True
    End With
    'myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
    With grCar.getGridObj
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 0)) = Trim(curCar_.id) Then
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
    tmpDet.VendorTruckID = selCar.id
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
