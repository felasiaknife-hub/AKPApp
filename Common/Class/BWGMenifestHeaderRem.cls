VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGMenifestHeaderRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RemID_ As String
Dim RemName_ As String
Dim colRem_ As Collection

Public Property Get RemID() As String
    RemID = RemID_
End Property
Public Property Let RemID(ByVal newVal As String)
    RemID_ = Trim(newVal)
End Property

Public Property Get RemName() As String
    RemName = RemName_
End Property
Public Property Let RemName(ByVal newVal As String)
    If Trim(RemName_) <> Trim(newVal) Then
        RemName_ = Trim(newVal)
    End If
End Property

Public Property Get Remset() As Collection
Dim tmpS As New BWGSearchManager
    If colRem_ Is Nothing Then Set colRem_ = tmpS.RemSearch("RemID='" & RemID_ & "'")
    Set Remset = colRem_
    Set tmpS = Nothing
End Property
'Public Property Let Remset(ByVal newVal As Collection)
'    Set Remset_ = newVal
'End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    RemID_ = selRec("RemID")
    RemName_ = Trim("" & selRec("RemName"))
End Sub

   
