VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTruckType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim typeDesc_ As String
Dim canBeTrailer_ As String
Dim isFactoryCar_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get TypeDesc() As String
    TypeDesc = typeDesc_
End Property
Public Property Let TypeDesc(ByVal newVal As String)
    If Trim(typeDesc_) <> Trim(newVal) Then
        typeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CanBeTrailer() As Boolean
    If canBeTrailer_ = "Y" Then
        CanBeTrailer = True
    Else
        CanBeTrailer = False
    End If
End Property
Public Property Let CanBeTrailer(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canBeTrailer_) <> Trim(tmpStr) Then
        canBeTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isFactoryCar() As Boolean
    If isFactoryCar_ = "Y" Then
        isFactoryCar = True
    Else
        isFactoryCar = False
    End If
End Property
Public Property Let isFactoryCar(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFactoryCar_) <> Trim(tmpStr) Then
        isFactoryCar_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbTruckType Where TruckTypeID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TruckTypeID") = id_
    End If
    tmpRec("TruckTypeDesc") = Trim(typeDesc_)
    tmpRec("CanBeTrailer") = Trim(canBeTrailer_)
    tmpRec("isFactoryCar") = Trim(isFactoryCar_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbTruckType Where TruckTypeID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("TruckTypeID")
    typeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    canBeTrailer_ = Trim("" & selRec("CanBeTrailer"))
    isFactoryCar_ = Trim("" & selRec("isFactoryCar"))
End Sub
