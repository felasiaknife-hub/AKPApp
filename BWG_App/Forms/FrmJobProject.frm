VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmJobProject 
   Caption         =   "บันทึกข้อมูลงานโปรเจค"
   ClientHeight    =   7440
   ClientLeft      =   5505
   ClientTop       =   2370
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   9195
   Begin VB.CommandButton cmdAttach 
      Caption         =   "..."
      Height          =   315
      Index           =   2
      Left            =   6600
      TabIndex        =   35
      Top             =   6210
      Width           =   435
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "..."
      Height          =   315
      Index           =   1
      Left            =   6600
      TabIndex        =   34
      Top             =   5850
      Width           =   435
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "ลบไฟล์"
      Height          =   315
      Index           =   2
      Left            =   8100
      TabIndex        =   33
      Top             =   6210
      Width           =   795
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "ลบไฟล์"
      Height          =   315
      Index           =   1
      Left            =   8100
      TabIndex        =   32
      Top             =   5850
      Width           =   795
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "เปิดไฟล์"
      Height          =   315
      Index           =   2
      Left            =   7140
      TabIndex        =   31
      Top             =   6210
      Width           =   795
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "เปิดไฟล์"
      Height          =   315
      Index           =   1
      Left            =   7140
      TabIndex        =   30
      Top             =   5850
      Width           =   795
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "เปิดไฟล์"
      Height          =   315
      Index           =   0
      Left            =   7140
      TabIndex        =   29
      Top             =   5490
      Width           =   795
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "ลบไฟล์"
      Height          =   315
      Index           =   0
      Left            =   8100
      TabIndex        =   28
      Top             =   5490
      Width           =   795
   End
   Begin VB.CommandButton cmdAttach 
      Caption         =   "..."
      Height          =   315
      Index           =   0
      Left            =   6600
      TabIndex        =   27
      Top             =   5490
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เสร็จงาน"
      Height          =   600
      Left            =   2745
      TabIndex        =   26
      Top             =   6705
      Width           =   1230
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้าง"
      Height          =   600
      Left            =   5265
      TabIndex        =   23
      Top             =   6705
      Width           =   1230
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   600
      Left            =   180
      TabIndex        =   19
      Top             =   6705
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "บันทึก"
      Height          =   600
      Left            =   4005
      TabIndex        =   18
      Top             =   6705
      Width           =   1230
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบ"
      Height          =   600
      Left            =   6525
      TabIndex        =   17
      Top             =   6705
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   600
      Left            =   7785
      TabIndex        =   16
      Top             =   6705
      Width           =   1230
   End
   Begin VB.TextBox txtProjectDetail 
      Height          =   2895
      Left            =   180
      TabIndex        =   15
      Top             =   2520
      Width           =   8835
   End
   Begin VB.TextBox txtProjectCost 
      Height          =   330
      Left            =   3780
      TabIndex        =   13
      Top             =   1935
      Width           =   1905
   End
   Begin VB.TextBox txtIncollectFee 
      Height          =   330
      Left            =   1125
      TabIndex        =   11
      Top             =   1935
      Width           =   1905
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   330
      Left            =   1665
      TabIndex        =   7
      Top             =   1440
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Format          =   80609281
      CurrentDate     =   44466
   End
   Begin VB.TextBox txtProjectName 
      Height          =   330
      Left            =   1035
      TabIndex        =   5
      Top             =   225
      Width           =   4650
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1035
      TabIndex        =   2
      Top             =   630
      Width           =   4665
   End
   Begin VB.TextBox txtSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      Top             =   1035
      Width           =   4665
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   330
      Left            =   4275
      TabIndex        =   9
      Top             =   1440
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      Format          =   80609281
      CurrentDate     =   44466
   End
   Begin MSComDlg.CommonDialog OpenDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "เลือกไฟล์"
   End
   Begin VB.Label lblAttach 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1335
      TabIndex        =   41
      Top             =   6270
      Width           =   5190
   End
   Begin VB.Label lblAttach 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1335
      TabIndex        =   40
      Top             =   5910
      Width           =   5190
   End
   Begin VB.Label Label12 
      Caption         =   "ไฟล์แนบที่ 3"
      Height          =   195
      Left            =   270
      TabIndex        =   39
      Top             =   6270
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "ไฟล์แนบที่ 2"
      Height          =   195
      Left            =   270
      TabIndex        =   38
      Top             =   5910
      Width           =   975
   End
   Begin VB.Label lblAttach 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1335
      TabIndex        =   37
      Top             =   5550
      Width           =   5190
   End
   Begin VB.Label Label10 
      Caption         =   "ไฟล์แนบที่ 1"
      Height          =   195
      Left            =   270
      TabIndex        =   36
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label lblCustID 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Height          =   195
      Left            =   7245
      TabIndex        =   25
      Top             =   2205
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblStaffID 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   195
      Left            =   5985
      TabIndex        =   24
      Top             =   2070
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblProjectID 
      Height          =   240
      Left            =   6120
      TabIndex        =   22
      Top             =   135
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
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
      Height          =   915
      Left            =   5805
      TabIndex        =   21
      Top             =   585
      Width           =   3255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "สถานะ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7155
      TabIndex        =   20
      Top             =   225
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายละเอียด"
      Height          =   195
      Left            =   270
      TabIndex        =   14
      Top             =   2295
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ต้นทุน"
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   1980
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายได้"
      Height          =   195
      Left            =   540
      TabIndex        =   10
      Top             =   1980
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เสร็จสิ้น"
      Height          =   195
      Left            =   3330
      TabIndex        =   8
      Top             =   1485
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เริ่มดำเนินงาน"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   1485
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อโปรเจค"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   270
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้แทนขาย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   675
      Width           =   375
   End
End
Attribute VB_Name = "FrmJobProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim curStaff_ As BWGCompanyStaff
Private Type AttacheFile
    FileName As String
    FullPath As String
End Type
Dim Attaches() As AttacheFile
Private Function AttacheFiles(ByVal Index As Integer, ByRef FileName As String, ByRef FullPath As String) As String
Dim StrFileName As String
Dim StrPath As String, StrFullPath As String
    OpenDialog.Filter = "PDF Files (*.pdf)|*.pdf|Pictures (*.jpg, *.bmp, *.gif, *.png| *.jpg; *.bmp; *.gif; *.png|" & _
    "World Document (*.doc)|*.doc|Zip (*.rar)|Work Sheet (*.xls)|*.xls|Visio (*.vsd)|*.vsd"
    OpenDialog.FileName = ""
    OpenDialog.ShowOpen
    If OpenDialog.FileName = "" Then
        AttacheFiles = ""
        Exit Function
    End If
    StrFullPath = OpenDialog.FileName
    StrFileName = OpenDialog.FileTitle
    StrPath = Mid(StrFullPath, 1, Len(StrFullPath) - Len(StrFileName))
    
    AttacheFiles = StrFileName
    FileName = StrFileName
    FullPath = StrFullPath
End Function

Private Sub cmdAttach_Click(Index As Integer)
Dim FileName As String
Dim FullPath As String
    If lblAttach(Index).Caption <> "" Then KillFile Index, lblCustID.Caption, lblProjectID.Caption
    AttacheFiles Index, FileName, FullPath
    lblAttach(Index).Caption = FileName
    Attaches(Index).FileName = FileName
    Attaches(Index).FullPath = FullPath
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCls_Click(Index As Integer)
    KillFile Index, lblCustID.Caption, lblProjectID.Caption
End Sub

Private Sub cmdDel_Click()
Dim i As Integer
    If MsgBox("ต้องการลบข้อมูลโปรเจคนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        DBConnExc "Delete tbJobProject WHERE ProjectID='" & lblProjectID.Caption & "' "
        ClrObj
        For i = 0 To 2
        KillFile i, lblCustID.Caption, lblProjectID.Caption
        Next
        MsgBox "ลบข้อมูลเรียบร้อยแล้ว", vbInformation
    End If
End Sub

Private Sub cmdNew_Click()
    ClrObj
End Sub

Private Sub cmdOpenFile_Click(Index As Integer)
Dim OpenFile As Double
        DefaultPath = "\\192.168.1.2\JobProject\"
    If lblAttach(Index).Caption <> "" Then
        ShellExecute 0, vbNullString, Attaches(Index).FullPath, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub CmdSave_Click()
Dim sDate As String
Dim EDate As String
    If Trim(lblCustID.Caption) = "" Then
        MsgBox "ข้อมูลลูกค้ายังไม่ได้รับการระบุ กรุณาตรวจสอบ", vbExclamation
        Exit Sub
    End If
    If Trim(lblStaffID.Caption) = "" Then
        MsgBox "ข้อมูลผู้แทนขายยังไม่ได้รับการระบุ กรุณาตรวจสอบ", vbExclamation
        Exit Sub
    End If
    If Trim(txtProjectName.Text) = "" Then
        MsgBox "ต้องระบุชื่อโปรเจคก่อน", vbExclamation
        Exit Sub
    End If
    If Trim(txtIncollectFee.Text) = "" Then
        MsgBox "ต้องระบุรายได้", vbExclamation
        Exit Sub
    End If
    If Trim(txtProjectCost.Text) = "" Then
        MsgBox "ต้องระบุต้นทุน", vbExclamation
        Exit Sub
    End If
    If Year(dtpStartDate.Value) <= "2550" Then
        sDate = CDbl(Year(dtpStartDate.Value) + 543)
        sDate = sDate & "/" & Format(Month(dtpStartDate.Value), "0#") & "/" & Format(Day(dtpStartDate.Value), "0#")
    End If
    If Year(dtpEndDate.Value) <= "2550" Then
        EDate = CDbl(Year(dtpEndDate.Value) + 543)
        EDate = EDate & "/" & Format(Month(dtpEndDate.Value), "0#") & "/" & Format(Day(dtpEndDate.Value), "0#")
    End If
            If Trim(lblProjectID.Caption) = "" Then
                DBConnExc "Insert into tbJobProject (ProjectID,CustomerID,ProjectName,SaleStaffID,StartDate,EndDate,IncollectFee,ProjectCost,ProjectDetail,ProjectStatus)" & _
                                        " Values ('" & GetGUID & "','" & lblCustID.Caption & "','" & Trim(txtProjectName.Text) & "','" & lblStaffID.Caption & "','" & sDate & "','" & EDate & "'," & Trim(txtIncollectFee.Text) & "," & Trim(txtProjectCost.Text) & ",'" & Trim(txtProjectDetail.Text) & "','W')"
            Else
                DBConnExc "Update tbJobProject Set CustomerID='" & Trim(lblCustID.Caption) & "',ProjectName='" & Trim(txtProjectName.Text) & "', SaleStaffID='" & Trim(lblStaffID.Caption) & "', " & _
                                        "StartDate='" & sDate & "', EndDate='" & EDate & "', IncollectFee=" & Trim(txtIncollectFee.Text) & ", ProjectCost=" & Trim(txtProjectCost.Text) & ", ProjectDetail= '" & Trim(txtProjectDetail.Text) & "' " & _
                                        " WHERE ProjectID='" & Trim(lblProjectID.Caption) & "' "
            End If
        If SaveAttache(lblCustID.Caption, lblProjectID.Caption) = False Then
            MsgBox "การบันทึกไฟล์แนบไม่สมบูรณ์", vbInformation, "ไฟล์แนบ"
        End If

        MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
    With frmSearchProject
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Command1_Click()
    If MsgBox("ยืนยันการเสร็จสิ้นโปรเจคนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        If Trim(lblProjectID.Caption) <> "" Then
            If lblStatus.Caption <> "เสร็จงาน" Then
                DBConnExc "Update tbJobProject Set ProjectStatus='E' WHERE ProjectID='" & Trim(lblProjectID.Caption) & "'"
                MsgBox "บันทึกข้อมูลการเสร็จงานเรียบร้อย", vbInformation
                SetdataProject (Trim(lblProjectID.Caption))
            Else
                Exit Sub
            End If
        Else
              MsgBox "กรุณาทำการค้นหาข้อมูลที่ต้องการก่อนกดปุ่มเสร็จงาน", vbInformation
              Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    ClrObj
    ReDim Attaches(3)
End Sub
Sub ClrObj()
    Set curCustomer_ = Nothing
    Set curStaff_ = Nothing
    txtCustomer.Text = ""
    txtSale.Text = ""
    txtIncollectFee.Text = ""
    txtProjectCost.Text = ""
    txtProjectDetail.Text = ""
    txtProjectName.Text = ""
    lblProjectID.Caption = ""
    lblStatus.Caption = ""
    dtpStartDate.Value = Now
    dtpEndDate.Value = Now
End Sub
Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If

End Sub
Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If

End Sub
Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
        lblCustID.Caption = .ID
    End With
End Sub
Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
        lblStaffID.Caption = .ID
    End With
End Sub
Public Sub SetdataProject(ByVal ProjectID As String)
        Dim Sql As String
        Dim Stat As String
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Sql = " select * from vw_jobProject  Where ProjectID = " & ProjectID & " "
        Set rs = GetRS(Sql, adOpenForwardOnly, adLockReadOnly)
    If rs.EOF = False Then
        lblProjectID.Caption = rs!ProjectID
        lblCustID.Caption = rs!CustomerID
        lblStaffID.Caption = rs!SaleStaffID
        txtProjectName.Text = rs!ProjectName
        txtCustomer.Text = rs!CompanyName
        txtSale.Text = rs!StaffFName & " " & rs!StaffLName
        dtpStartDate.Value = rs!StartDate
        dtpEndDate.Value = rs!EndDate
        txtIncollectFee.Text = rs!InCollectFee
        txtProjectCost.Text = rs!ProjectCost
        txtProjectDetail.Text = rs!ProjectDetail
        If rs!ProjectStatus = "W" Then
            Stat = "รอดำเนินการ"
        ElseIf rs!ProjectStatus = "Q" Then
            Stat = "เสนอราคา"
        ElseIf rs!ProjectStatus = "P" Then
            Stat = "กำลังดำเนินการ"
        ElseIf rs!ProjectStatus = "E" Then
            Stat = "เสร็จงาน"
        End If
        lblStatus.Caption = Stat
        
        LoadAttacheFile lblCustID.Caption, lblProjectID.Caption
        
    End If
End Sub

Private Function SaveAttache(ByVal CustomerID As String, ByVal ProjectID As String) As Boolean
On Error GoTo ErrD
Dim i As Integer
Dim ServerPath As String
Dim SourceFile As String
    SaveAttache = True
        DefaultPath = "\\192.168.1.2\JobProject\"
        ServerPath = DefaultPath & CustomerID & "\"
    '======= Check Path ===================
    If Dir(ServerPath, vbDirectory) = "" Then
      MkDir ServerPath
    End If
    ServerPath = ServerPath & ProjectID & "\"
    If Dir(ServerPath, vbDirectory) = "" Then
      MkDir ServerPath
    End If
    
    For i = 0 To 2
        If Attaches(i).FileName <> "" Then
            FileCopy Attaches(i).FullPath, ServerPath & Attaches(i).FileName
        End If
    Next
    Exit Function
ErrD:
    SaveAttache = False
    Resume Next
End Function

Private Function KillFile(ByVal Index As Integer, ByVal CustomerID As String, ByVal ProjectID As String) As Boolean
On Error GoTo ErrD
Dim FileName As String
    KillFile = True
    FileName = DefaultPath & CustomerID & "\" & ProjectID & "\" & lblAttach(Index).Caption
    If Dir(FileName, vbNormal) <> "" Then
        Kill FileName
        lblAttach(Index).Caption = ""
    End If
    Exit Function
ErrD:
    KillFile = False
End Function

Private Sub LoadAttacheFile(ByVal CustomerID As String, ByVal ProjectID As String)
On Error GoTo ErrD
Dim FSO As New FileSystemObject
Dim Filex As File
Dim FileObject As Files
Dim FolderObject As Folder
Dim ServerFilePath As String
Dim i As Integer, j As Integer, x As Integer, Files As String
Dim Extension As String
    DefaultPath = "\\192.168.1.2\JobProject\"
    ServerFilePath = DefaultPath & CustomerID & "\" & ProjectID
    If FSO.FolderExists(ServerFilePath) = True Then
        Set FolderObject = FSO.GetFolder(ServerFilePath)
        Set FileObject = FolderObject.Files
        i = FileObject.Count
        x = 0
        ReDim Attaches(4)
        For Each Filex In FileObject
            For j = Len(Filex.Name) To 1 Step -1
                Extension = Right(Filex.Name, Len(Filex.Name) - j)
                If Mid(Filex.Name, j, 1) = "." Then Exit For
            Next
            If (LCase(Extension) <> "db") And (LCase(Extension) <> "dll") And (LCase(Extension) <> "exe") And (LCase(Extension) <> "com") And (LCase(Extension) <> "bat") Then
                Attaches(x).FileName = Filex.Name
                Attaches(x).FullPath = ServerFilePath & "\" & Filex.Name
                lblAttach(x).Caption = Filex.Name
                x = x + 1
            End If
            If x > 3 Then Exit For
        Next Filex
    End If
    
    Exit Sub
ErrD:
    MsgBox err.Description
End Sub

