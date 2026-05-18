VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBoxOwners 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Box Owners"
   ClientHeight    =   5250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBoxOwnersDesc 
      Height          =   315
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   2
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox txtBoxOwnersCode 
      Height          =   315
      Left            =   120
      MaxLength       =   3
      TabIndex        =   1
      Top             =   4200
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   4740
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   5100
      TabIndex        =   5
      Top             =   4740
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   3480
      TabIndex        =   4
      Top             =   4740
      Width           =   1395
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Owners Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Owners Description"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmBoxOwners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    txtBoxOwnersCode.Text = ""
    txtBoxOwnersDesc.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim tSql As String
    GetDBConnection DBConStock
    tSql = "SELECT BoxOwnersCode, BoxOwnersDesc FROM Box_BoxOwners WHERE BoxOwnersCode = '" & Trim(txtBoxOwnersCode.Text) & "'"
    Set RS = New ADODB.Recordset
    RS.open tSql, DBConn, adOpenDynamic, adLockOptimistic
    If RS.EOF Then
        RS.AddNew
        RS!BoxOwnersCode = Trim(txtBoxOwnersCode.Text)
        RS!BoxOwnersDesc = Trim(txtBoxOwnersDesc.Text)
    Else
        RS!BoxOwnersDesc = Trim(txtBoxOwnersDesc.Text)
    End If
    RS.Update
    Set RS = Nothing
    InitialData
    MsgBox "Saving completed", vbInformation
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    InitialData
End Sub

Private Sub InitialData()
Dim RS As ADODB.Recordset
Dim tSql As String, tmpList As ListItem
    GetDBConnection DBConStock
    tSql = "SELECT BoxOwnersCode, BoxOwnersDesc FROM Box_BoxOwners ORDER BY BoxOwnersCode"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    ListView1.ListItems.Clear
    Do While Not RS.EOF
        Set tmpList = ListView1.ListItems.Add(, "'" & RS!BoxOwnersCode & "'", RS!BoxOwnersCode)
        tmpList.SubItems(1) = RS!BoxOwnersDesc
        RS.MoveNext
    Loop
    Set tmpList = Nothing
    Set RS = Nothing
    If ListView1.ListItems.Count > 0 Then ListView1_ItemClick ListView1.SelectedItem
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtBoxOwnersCode.Text = Item.Text
    txtBoxOwnersDesc.Text = Item.SubItems(1)
End Sub
