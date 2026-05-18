VERSION 5.00
Begin VB.Form frmCustOtherRemark 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frCustOtherRemark.frx":0000
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   570
      Width           =   6525
   End
   Begin VB.TextBox txtOrdPos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8370
      TabIndex        =   2
      Top             =   570
      Width           =   465
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1350
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1350
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1350
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1350
      Width           =   1005
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   1230
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อบิล/สถานที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   630
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลำดับที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   7830
      TabIndex        =   8
      Top             =   630
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   510
      TabIndex        =   7
      Top             =   210
      Width           =   570
   End
End
Attribute VB_Name = "frmCustOtherRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myRemark_ As BWGCustOrderRemark
Dim curCustomer_ As BWGCustomer
Dim callerForm_ As Form

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลหมายเหตุทั่วไปของลูกค้ารายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myRemark_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewRemark
            txtRemark.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myRemark_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myRemark_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewRemark
    txtRemark.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If

    If Trim(txtRemark.Text) = "" Then
        MsgBox "กรุณาใส่หมายเหตุ", vbExclamation
        txtRemark.SetFocus
        Exit Sub
    End If

    If Trim(txtOrdPos.Text) = "" Then
        MsgBox "กรุณาระบุลำดับที่", vbExclamation
        txtOrdPos.SetFocus
        Exit Sub
    End If

    If IsNumeric(Trim(txtOrdPos.Text)) = False Then
        MsgBox "ลำดับที่ไม่ถูกต้อง", vbExclamation
        txtOrdPos.SetFocus
        Exit Sub
    End If

    If myRemark_.isChanged Then
        myRemark_.Save
    End If
    If myRemark_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
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
        cmdClose.Picture = .picExit.Picture
    End With
    CreateNewRemark
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtRemark.Text = ""
    txtOrdPos.Text = ""
End Sub

Private Sub CreateNewRemark()
    Set myRemark_ = Nothing
    Set myRemark_ = New BWGCustOrderRemark
    With myRemark_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curCustomer_ Is Nothing Then .CustomerID = curCustomer_.ID
        .isChanged = False
    End With
End Sub

Public Sub SetDataBillingAddr(selSite As BWGCustOrderRemark)
    Set myRemark_ = Nothing
    Set myRemark_ = selSite
    If myRemark_ Is Nothing Then Exit Sub
    With myRemark_
        txtRemark.Text = .RemDesc
        txtOrdPos.Text = .OrdPos
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myRemark_ Is Nothing Then Exit Sub
    If myRemark_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myRemark_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtOrdPos_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtOrdPos.Text)) Then
        myRemark_.OrdPos = Trim(txtOrdPos.Text)
    Else
        myRemark_.OrdPos = 0
    End If
End Sub

Private Sub txtRemark_Validate(Cancel As Boolean)
    myRemark_.RemDesc = Trim(txtRemark.Text)
End Sub

Public Sub SetCallerForm(selFrm As Form, selCust As BWGCustomer)
    Set callerForm_ = selFrm
    Set curCustomer_ = selCust
    txtCustomer.Text = selCust.CustomerName
    If Not myRemark_ Is Nothing Then myRemark_.CustomerID = curCustomer_.ID
End Sub
