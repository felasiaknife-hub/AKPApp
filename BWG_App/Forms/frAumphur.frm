VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAumphur 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F002"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTumbol 
      Caption         =   "ข้อมูลตำบล"
      Height          =   405
      Left            =   5880
      TabIndex        =   11
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1770
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   2130
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1770
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1770
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1770
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1770
      Width           =   1005
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1770
      Width           =   1005
   End
   Begin VB.ComboBox cboProvince 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2835
   End
   Begin VB.TextBox txtAumphur 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4530
      MaxLength       =   255
      TabIndex        =   1
      Top             =   840
      Width           =   2595
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Line Line2 
      X1              =   -960
      X2              =   7260
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อจังหวัด"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5910
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   6060
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอ"
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
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่ออำเภอ"
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   900
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frAumphur.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7320
   End
End
Attribute VB_Name = "frmAumphur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myAumphur_ As IMIAumphur
Dim colProvince_ As Collection
Dim callerForm_ As Form

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub cboProvince_Validate(Cancel As Boolean)
    If myAumphur_ Is Nothing Then
        CreateNewAumphur
    End If
    If cboProvince.ListIndex > -1 Then
        myAumphur_.ProvinceID = colProvince_(cboProvince.ListIndex + 1).ID
    Else
        myAumphur_.ProvinceID = ""
    End If
End Sub

Private Sub cmdPrint_Click()
Dim selResult As Integer
    If myAumphur_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myAumphur_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    With rpt1
        .Connect = "dsn=" & DSN_NAME
        .ReportFileName = App.Path & "\reports\Aumphur.rpt"
        .Destination = crptToWindow
        .WindowShowPrintSetupBtn = True
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
        .PrintReport
    End With

End Sub

Private Sub cmdTumbol_Click()
Dim selResult As Integer
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmTumbol.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    If Trim(txtAumphur.Text) = "" Then
        MsgBox "กรุณาใส่ชื่ออำเภอ", vbExclamation
        txtAumphur.SetFocus
        Exit Sub
    End If
    If myAumphur_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myAumphur_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmTumbol
    frmTumbol.SetCurAumphur myAumphur_
    frmTumbol.Show 1
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdPrint.Picture = .picPrint.Picture
        'cmdPrintInv.Picture = .picPrint.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    PopulateProvince
    CreateNewAumphur
End Sub

Private Sub txtAumphur_Validate(Cancel As Boolean)
    If myAumphur_ Is Nothing Then
        CreateNewAumphur
    End If
    myAumphur_.Name = Trim(txtAumphur.Text)
End Sub

Private Sub PopulateProvince()
Dim x%
Dim tmpS As New BWGSearchManager
    Set colProvince_ = tmpS.ProvinceSearch
    cboProvince.Clear
    For x = 1 To colProvince_.Count
        cboProvince.AddItem colProvince_(x).ProvinceName
    Next
    Set tmpS = Nothing
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ต้องการลบข้อมูลอำเภอนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myAumphur_.Delete
        ClearScreen
        CreateNewAumphur
        cboProvince.SetFocus
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myAumphur_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myAumphur_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewAumphur
    cboProvince.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
    If myAumphur_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myAumphur_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    Load frmAumphurSearch
    With frmAumphurSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub CmdSave_Click()
Dim tmpS As BWGSearchManager
Dim tmpCol As Collection
    If Trim(cboProvince.Text) = "" Then
        MsgBox "กรุณาเลือกจังหวัด", vbExclamation
        cboProvince.SetFocus
        Exit Sub
    End If
    If Trim(txtAumphur.Text) = "" Then
        MsgBox "กรุณาใส่ชื่ออำเภอ", vbExclamation
        txtAumphur.SetFocus
        Exit Sub
    End If
    Set tmpS = New BWGSearchManager
    Set tmpCol = tmpS.AumphurSearch("AumphurName='" & Trim(txtAumphur.Text) & "' And ProvinceID='" & myAumphur_.ProvinceID & "'")
    If tmpCol.Count > 0 Then
        If Trim(tmpCol(1).ID) <> Trim(myAumphur_.ID) Then
            MsgBox "มีชื่ออำเภอ " & txtAumphur.Text & " ในจังหวัด " & cboProvince.Text & " แล้ว ไม่สามารถเพิ่มซ้ำได้", vbExclamation
            Exit Sub
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End If
    If myAumphur_.isChanged Then
        myAumphur_.Save
    End If
    If myAumphur_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myAumphur_ Is Nothing Then Exit Sub
    If myAumphur_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myAumphur_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub CreateNewAumphur()
    Set myAumphur_ = Nothing
    Set myAumphur_ = New IMIAumphur
    With myAumphur_
        .ID = getNewProductID("AP")
        If cboProvince.ListIndex > -1 Then
            .ProvinceID = colProvince_(cboProvince.ListIndex + 1).ID
        End If
        .isChanged = False
    End With
End Sub

Private Sub ClearScreen()
    cboProvince.ListIndex = -1
    txtAumphur.Text = ""
End Sub

Public Sub SetData(selPratunam As IMIAumphur)
    Set myAumphur_ = Nothing
    Set myAumphur_ = selPratunam
    If myAumphur_ Is Nothing Then Exit Sub
    With myAumphur_
        txtAumphur.Text = .Name
        If isExist(colProvince_, "'" & .ProvinceID & "'") Then
            SetTextToCombo colProvince_("'" & .ProvinceID & "'").ProvinceName, cboProvince
        End If
    End With
End Sub
