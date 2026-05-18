VERSION 5.00
Begin VB.Form frmLABTemplate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F050"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   30
      TabIndex        =   12
      Top             =   840
      Width           =   11325
      Begin VB.TextBox txtTemplateNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox txtTemplateName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5790
         TabIndex        =   3
         Top             =   780
         Width           =   4515
      End
      Begin VB.ComboBox cboTestType 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   2445
      End
      Begin VB.TextBox txtCreateDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5790
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อ Template"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4770
         TabIndex        =   16
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Template No."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   750
         TabIndex        =   15
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5370
         TabIndex        =   14
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทการทดสอบ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ข้อมูล LABTemplate"
      Height          =   4665
      Left            =   0
      TabIndex        =   11
      Top             =   2190
      Width           =   11355
      Begin WasteManagment.ctlGrid grLABDetail 
         Height          =   4335
         Left            =   60
         TabIndex        =   4
         Top             =   270
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   7646
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6900
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6900
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6900
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6900
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6900
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
      Caption         =   "สร้างแบบฟอร์มการทดสอบ (LAB Template)"
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
      TabIndex        =   10
      Top             =   60
      Width           =   3810
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frLABTemplate.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmLABTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myLABTemplateHead_ As BWGLABTemplateHeader
Dim colTestType_ As Collection
Dim colLABMethod_ As Collection
Dim colLABParam_ As Collection

Private Sub SetMSFlexGrid()
   
    grLABDetail.ColInputType(1) = ComboBoxEnum
    grLABDetail.ColInputType(2) = ComboBoxEnum
    grLABDetail.ColInputType(3) = ComboBoxEnum
    grLABDetail.ColInputType(5) = ComboBoxEnum
    grLABDetail.ColInputType(6) = ComboBoxEnum

    With grLABDetail.getGridObj
        .Cols = 8
        .ColWidth(0) = 0
        .ColWidth(1) = 1845
        .ColWidth(2) = 1200
        .ColWidth(3) = 1500
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 1500
        .ColWidth(7) = 2520
        .Rows = 100
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 1
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
    
        .TextMatrix(0, 1) = "Parameter"
        .TextMatrix(0, 2) = "Unit"
        .TextMatrix(0, 3) = "Method"
        .TextMatrix(0, 4) = "Standard"
        .TextMatrix(0, 5) = "ประเภทผลลัพธ์"
        .TextMatrix(0, 6) = "การเปรียบเทียบ"
        .TextMatrix(0, 7) = "คำอธิบาย"
    End With
End Sub

Private Sub cboTestType_Validate(Cancel As Boolean)
    If cboTestType.ListIndex > -1 Then
        myLABTemplateHead_.TestTypeID = colTestType_(cboTestType.ListIndex + 1).ID
        myLABTemplateHead_.TestTypeName = cboTestType.Text
    Else
        myLABTemplateHead_.TestTypeID = ""
        myLABTemplateHead_.TestTypeName = ""
    End If
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myLABTemplateHead_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            cboTestType.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myLABTemplateHead_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLABTemplateHead_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    ClearScreen
    CreateNewLABTemplateHeader
    
    PopulateLABMethod
    PopulateLABParam
    PopulateLABTestType
    cboTestType.SetFocus
End Sub
Private Sub ClearScreen()
    txtCreateDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    txtTemplateNo.Text = "== AUTO =="
    txtTemplateName.Text = ""
    cboTestType.ListIndex = -1
    grLABDetail.ClearAllData
    grLABDetail.getGridObj.Rows = 100
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim x%
Dim tmpDet As BWGLABTemplateDetail
    If Trim(txtTemplateName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อTemplate", vbExclamation
        txtTemplateName.SetFocus
        Exit Sub
    End If
    If Trim(cboTestType.Text) = "" Then
        MsgBox "กรุณาระบุประเภทการทดสอบ", vbExclamation
        cboTestType.SetFocus
        Exit Sub
    End If

    For x = 1 To myLABTemplateHead_.LABTemplateDet.Count
        Set tmpDet = myLABTemplateHead_.LABTemplateDet(x)
        If Trim(tmpDet.ParameterID) = "" Then
            MsgBox "กรุณาระบุข้อมูล Parameter ให้ครบถ้วน", vbExclamation
            Set tmpDet = Nothing
            
            Exit Sub
        End If
        If Trim(tmpDet.MethodID) = "" Then
            MsgBox "กรุณาระบุข้อมูล Method ให้ครบถ้วน", vbExclamation
            Set tmpDet = Nothing
            Exit Sub
        End If
        
    Next
    Set tmpDet = Nothing
    If myLABTemplateHead_.isChanged Then
        myLABTemplateHead_.Save
        txtTemplateNo.Text = myLABTemplateHead_.TemplateNo
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    
    If myLABTemplateHead_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLABTemplateHead_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmLABTemplateSearch
    With frmLABTemplateSearch
        .SetCallerForm Me
        .Show 1
    End With
    txtTemplateName.SetFocus
    PopulateLABMethod
    PopulateLABParam
    PopulateLABTestType
    PopulateUnit
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
    txtCreateDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    Set myLABTemplateHead_ = Nothing
    Set myLABTemplateHead_ = New BWGLABTemplateHeader
     PopulateLABTestType
     PopulateLABParam
     PopulateLABMethod
     PopulateUnit
     CreateNewLABTemplateHeader
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grLABDetail" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub CreateNewLABTemplateHeader()
Set myLABTemplateHead_ = Nothing
Set myLABTemplateHead_ = New BWGLABTemplateHeader
    With myLABTemplateHead_
        .ID = GetGUID
        .CreatedDate = TodayDate
        If cboTestType.ListIndex > 0 Then
            myLABTemplateHead_.TestTypeID = colTestType_(cboTestType.ListIndex).ID
        Else
            myLABTemplateHead_.TestTypeID = ""
        End If
        .isChanged = False
    End With
    SetEnabledScreen
End Sub
Private Sub SetEnabledScreen()
Dim eFlag As Boolean
    eFlag = Not myLABTemplateHead_.isChanged
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

Private Sub PopulateLABTestType()
Dim x%
Dim tmpS As New BWGSearchManager
    Set colTestType_ = tmpS.TestTypeLABSearch
    cboTestType.Clear
    For x = 1 To colTestType_.Count
        cboTestType.AddItem colTestType_(x).TestTypeName
    Next
    'cboTestType.ListIndex = 0
    Set tmpS = Nothing
End Sub

Public Sub SetDataLABTemplate(selCust As BWGLABTemplateHeader, Optional isAssigned As Boolean = True)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
        
        Set myLABTemplateHead_ = selCust
        With myLABTemplateHead_
            txtCreateDate.Text = FormatYMD_to_DMY(.CreatedDate, "/", "/")
            txtTemplateNo.Text = .TemplateNo
            txtTemplateName.Text = .TemplateName
            cboTestType.Text = .TestTypeName
        End With
        ShowLABTemplateDetail
End Sub

Private Sub ShowLABTemplateDetail()
Dim x%, Y%
Dim tmpCol As Collection
Dim tmpDet As BWGLABTemplateDetail
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    With grLABDetail.getGridObj
        grLABDetail.ClearAllData

        Set tmpCol = myLABTemplateHead_.LABTemplateDet
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.ID
            .row = x + rowCount
            '.col = 1
            .CellForeColor = forColor
            .TextMatrix(x + rowCount, 1) = tmpDet.ParameterName
            .TextMatrix(x + rowCount, 2) = tmpDet.LabUnitName
            .TextMatrix(x + rowCount, 3) = tmpDet.MethodName
            .TextMatrix(x + rowCount, 4) = tmpDet.StandardValue
            If tmpDet.hasStandard = True Then
                .TextMatrix(x + rowCount, 5) = "ตัวเลข"
            ElseIf tmpDet.hasStandard = False Then
                .TextMatrix(x + rowCount, 5) = "ตัวหนังสือ"
            End If
            .TextMatrix(x + rowCount, 6) = tmpDet.CompareType
            .TextMatrix(x + rowCount, 7) = tmpDet.CommentDesc
        Next
        '.row = 1
        '.col = 1
    End With
    Set tmpCol = Nothing
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myLABTemplateHead_ Is Nothing Then Exit Sub
    If myLABTemplateHead_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myLABTemplateHead_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub


Private Sub grLABDetail_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 4 And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลข", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grLABDetail_GotFocus()
    If grLABDetail.getGridObj.col = 0 Then grLABDetail.getGridObj.col = 1
End Sub

Private Sub grLABDetail_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grLABDetail.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการทดสอบนี้หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = myLABTemplateHead_.LABTemplateDet
                selCol.Remove tmpKey
                myLABTemplateHead_.isChanged = True
            End If
        End If
    End With

End Sub

Private Sub grLABDetail_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpS As BWGSearchManager
Dim tmpCbo As ComboBox
Dim tmpKey As String
    
    With grLABDetail.getGridObj
        If Trim(.TextMatrix(row - 1, 1)) = "" And col > 1 Then
            grLABDetail.ColEnabled(col) = False
            Exit Sub
        Else
           If col = 1 Or col = 2 Or col = 3 Or col = 4 Or col = 5 Or col = 6 Or col = 7 Then
                grLABDetail.ColEnabled(col) = True
                
            Else
                grLABDetail.ColEnabled(col) = False
                Exit Sub
            End If
        End If
    End With

    tmpKey = "'" & grLABDetail.getGridObj.TextMatrix(row, 0) & "'"
    'If isExist(myLABTemplateHead_.LABTemplateDet, tmpKey) Then
        If (col = 1) And Not colLABParam_ Is Nothing Then
            Set tmpCbo = grLABDetail.getColObject(1)
            tmpCbo.Clear
            For x = 1 To colLABParam_.Count
                tmpCbo.AddItem colLABParam_(x).ParameterName
            Next
            Set tmpCbo = Nothing
        ElseIf (col = 2) And Not colUnit_ Is Nothing Then
            Set tmpCbo = grLABDetail.getColObject(2)
            tmpCbo.Clear
            For x = 1 To colUnit_.Count
                tmpCbo.AddItem colUnit_(x).UnitName
            Next
            Set tmpCbo = Nothing
        ElseIf (col = 3) And Not colLABMethod_ Is Nothing Then
            Set tmpCbo = grLABDetail.getColObject(3)
            tmpCbo.Clear
            For x = 1 To colLABMethod_.Count
                tmpCbo.AddItem colLABMethod_(x).MethodName
            Next
            Set tmpCbo = Nothing
        ElseIf (col = 5) Then
            Set tmpCbo = grLABDetail.getColObject(5)
            tmpCbo.Clear
            tmpCbo.AddItem "ตัวเลข"
            tmpCbo.AddItem "ตัวหนังสือ"
            Set tmpCbo = Nothing
        ElseIf (col = 6) Then
            Set tmpCbo = grLABDetail.getColObject(6)
            tmpCbo.Clear
            tmpCbo.AddItem ">"
            tmpCbo.AddItem "<"
            tmpCbo.AddItem "="
            tmpCbo.AddItem "<="
            tmpCbo.AddItem ">="
            tmpCbo.AddItem "<>"
            Set tmpCbo = Nothing
        End If
    'End If

End Sub

Private Sub grLABDetail_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGLABTemplateDetail
Dim tmpKey As String
Dim tmpVal As String
Dim tmpRow%

    tmpKey = Trim(grLABDetail.getGridObj.TextMatrix(row, 0))
    If col = 1 And Trim(grLABDetail.getGridObj.TextMatrix(row, col)) = "" Then Exit Sub
    If isExist(myLABTemplateHead_.LABTemplateDet, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGLABTemplateDetail
        tmpDet.ID = GetGUID
        tmpDet.TemplateID = myLABTemplateHead_.ID
        tmpDet.ItemNo = row
        grLABDetail.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        myLABTemplateHead_.LABTemplateDet.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myLABTemplateHead_.LABTemplateDet("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grLABDetail.getGridObj.TextMatrix(row, col))

    With tmpDet
        Select Case col
            Case 1
                If Trim(tmpVal) <> "" Then
                    .ParameterID = colLABParam_(grLABDetail.getColObject(1).ListIndex + 1).ID
                    .ParameterName = Trim(tmpVal)
                Else
                    .ParameterID = ""
                    .ParameterName = ""
                End If
            Case 2
                If Trim(tmpVal) <> "" Then
                    .LabUnitID = colUnit_(grLABDetail.getColObject(2).ListIndex + 1).ID
                    .LabUnitName = Trim(tmpVal)
                Else
                    .LabUnitID = ""
                    .LabUnitName = ""
                End If
            Case 3
                If Trim(tmpVal) <> "" Then
                    .MethodID = colLABMethod_(grLABDetail.getColObject(3).ListIndex + 1).ID
                    .MethodName = Trim(tmpVal)
                Else
                    .MethodID = ""
                    .MethodName = ""
                End If
            Case 4
                    .StandardValue = Trim(tmpVal)
                    If Trim(tmpVal) <> "" Then
                        .hasStandard = True
                    Else
                        .hasStandard = False
                    End If
            Case 5
                If Trim(tmpVal) <> "" Then
                   If Trim(tmpVal) = "ตัวเลข" Then
                        .ResultType = "N"
                   Else
                        .ResultType = "S"
                   End If
                End If
            Case 6
                .CompareType = Trim(tmpVal)
            Case 7
                .CommentDesc = Trim(tmpVal)
        End Select

        myLABTemplateHead_.isChanged = True
    End With
    Set tmpDet = Nothing

End Sub

Private Sub txtCreateDate_Validate(Cancel As Boolean)
    'myLABTemplateHead_.CreatedDate = txtCreateDate.Text
End Sub

Private Sub txtTemplateName_Validate(Cancel As Boolean)
    myLABTemplateHead_.TemplateName = txtTemplateName.Text
End Sub
