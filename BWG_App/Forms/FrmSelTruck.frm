VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSelTruck 
   Caption         =   "เลือกรถ"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5190
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5190
      Width           =   1035
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1095
      MaxLength       =   1000
      TabIndex        =   1
      Top             =   870
      Width           =   5340
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   795
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5190
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3825
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   6747
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ประเภทรถ"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ราคา"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ประเภทรถ"
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   4
      Top             =   930
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายการรถสำหรับเลือกรถ"
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
      Left            =   225
      TabIndex        =   3
      Top             =   0
      Width           =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   5790
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5640
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "FrmSelTruck.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "FrmSelTruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colRemDesc_ As Collection
Dim colOriSelCar_ As Collection
Dim curRemDesc_ As BWGTruckType
Dim curRole_ As AWSRoleType
Dim callerFrm_ As Form
Dim tumbolID_ As String
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim colMergeCollection As Collection
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim x%, i As Integer
Dim tmpKey As String
 If colOriSelCar_ Is Nothing Then Set colOriSelCar_ = New Collection
    For x = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(x).id & "'")
        If isExist(colOriSelCar_, tmpKey) Then
            callerFrm_.RemoveCar colItemDel_(x).id
        End If
    Next
    For x = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(x).id & "'")
        If isExist(colOriSelCar_, tmpKey) = False Then
            colOriSelCar_.Add colItemAdd_(x), tmpKey
            callerFrm_.AddCartem colItemAdd_(x)
        End If
    Next
    Unload Me

End Sub

Private Sub cmdSearch_Click()
Dim x%, i As Integer
Dim searchM As New BWGSearchManager
    With frmPicture
        cmdClose.Picture = .picExit.Picture
    End With
    Set colRemDesc_ = New Collection
    If tumbolID_ <> "" Then
    Set colRemDesc_ = searchM.TruckRateSearch("TumbolID= '" & tumbolID_ & "'")
    Else
    Set colRemDesc_ = searchM.TruckRateSearch
    End If

    Set colMergeCollection = New Collection
      For i = 1 To colRemDesc_.Count
        colMergeCollection.Add colRemDesc_.item(i)
    Next
    ShowTruckList

End Sub

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerFrm_ = frm
    Set colOriSelCar_ = frm.getCartems
    cmdSearch_Click
    
End Sub
Public Sub SetTumbolID(id_ As String)
    tumbolID_ = id_
End Sub
Private Sub ShowTruckList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGTruckRate

    If colRemDesc_ Is Nothing Then Exit Sub
    lvResult.ListItems.Clear
    For x = 1 To colRemDesc_.Count
        Set tmpItem = colRemDesc_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.id & "'", tmpItem.TypeDesc)
        tmpList.SubItems(1) = Format(tmpItem.Rate, "#,###.00")
        tmpList.Checked = isExist(colOriSelCar_, tmpList.Key)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub



Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked = True Then
        colItemAdd_.Add colMergeCollection(item.Index), item.Key
         If isExist(colItemDel_, item.Key) Then
            colItemDel_.Remove item.Key
        End If
     Else
         If isExist(colItemDel_, item.Key) = False Then
            colItemDel_.Add colMergeCollection(item.Index), item.Key
        End If
        If isExist(colItemAdd_, item.Key) Then
            colItemAdd_.Remove item.Key
        End If
     End If
End Sub
Private Function CheckDiffType(ByVal wastedataID, wType As Boolean) As Boolean
    Dim i As Integer, tmpType As Boolean
    CheckDiffType = False: tmpType = False
    For i = 1 To colItemAdd_.Count
        tmpType = colItemAdd_.item(i).isQuotationReq
        If wType <> tmpType Then
            CheckDiffType = True
            Exit For
        Else
            CheckDiffType = False
        End If
    Next i
End Function
