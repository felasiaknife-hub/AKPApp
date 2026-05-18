VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserGroupSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F141"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   90
      Picture         =   "frUserGroupSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5100
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7590
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   1110
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   6906
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "กลุ่มผู้ใช้ระบบ"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "คำอธิบาย"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "สถานะ"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   6030
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   6180
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหากลุ่มผู้ใช้ระบบ"
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
      TabIndex        =   4
      Top             =   60
      Width           =   1725
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -120
      Top             =   750
      Width           =   12285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ผลการค้นหา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   -30
      TabIndex        =   3
      Top             =   840
      Width           =   8640
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frUserGroupSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10050
   End
End
Attribute VB_Name = "frmUserGroupSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
Dim selKey As String
Dim x%
    If lvResult.ListItems.Count > 0 Then
        If Not lvResult.SelectedItem Is Nothing Then
            selKey = lvResult.SelectedItem.Key
            If isExist(ColUserGroup, selKey) Then
                callerForm_.SetData ColUserGroup(selKey)
                Unload Me
            Else
                MsgBox "ไม่พบข้อมูลบริษัทที่ต้องการค้นหา", vbExclamation
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim x%
Dim tmpGroup As AWSUserGroup
Dim tmpList As ListItem
    With frmPicture
        cmdClose.Picture = .picExit.Picture
    End With

    PopulateUserGroup
    For x = 1 To ColUserGroup.Count
        Set tmpGroup = ColUserGroup(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpGroup.ID & "'", tmpGroup.Name)
        tmpList.SubItems(1) = tmpGroup.Description
        If tmpGroup.isActive Then
            tmpList.SubItems(2) = "พร้อมใช้งาน"
        Else
            tmpList.SubItems(2) = "ระงับใช้งานชั่วคราว"
        End If
        Set tmpList = Nothing
        Set tmpGroup = Nothing
    Next

End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub

Private Sub txtCriteria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub
