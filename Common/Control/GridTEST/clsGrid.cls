VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents fg As MSFlexGrid
Attribute fg.VB_VarHelpID = -1
Dim isRaiseBeforeUpdate_ As Boolean
Dim isRaiseUpdate_ As Boolean
Dim oldString As String
Dim isWait_ As Boolean
Dim colEnabled_(20) As Integer
Dim colType_(20) As ColType
Dim cancelUpdate_ As Boolean
Dim autoDel_ As Boolean
Dim gridMoveType_ As Integer

Public Property Set getGridControl(fgControl As MSFlexGrid)
    Set fg = fgControl
End Property

Public Function getColObject(selCol As Integer) As Object
    If colType_(selCol) = ComboBoxEnum Then
        Set getColObject = cboInput
    ElseIf colType_(selCol) = DateTimeEnum Then
        Set getColObject = dtInput
    ElseIf colType_(selCol) = TextBoxEnum Then
        Set getColObject = txtInput
    End If
    If Err.Description <> "" Then MsgBox Err.Description
End Function

Public Property Get ColEnabled(col As Integer) As Boolean
    ColEnabled(col) = CBool(colEnabled_(col))
End Property

Public Property Let ColEnabled(col As Integer, ByVal newVal As Boolean)
    colEnabled_(col) = CInt(newVal)
End Property

Public Property Get ColInputType(selCol As Integer) As ColType
    ColInputType = colType_(selCol)
End Property

Public Property Let ColInputType(selCol As Integer, ByVal newVal As ColType)
    colType_(selCol) = newVal
End Property

Public Property Get AutoDelete() As Boolean
    AutoDelete = autoDel_
End Property
Public Property Let AutoDelete(ByVal newVal As Boolean)
    autoDel_ = newVal
End Property

Public Sub GridMoveType(moveType As Integer)
    gridMoveType_ = moveType
End Sub

Public Sub WaitLostFocust(isWait As Boolean)
    isWait_ = isWait
End Sub

Private Sub SetTextToCombo(selTxt As String, selCbo As ComboBox)
On Error GoTo setErr:
    selCbo.Text = Trim(selTxt)
    Exit Sub
setErr:
    selCbo.ListIndex = -1
    Exit Sub
End Sub

Private Function AddZero(str As String) As String
    AddZero = Format(str, "00")
End Function

Private Function FormatYMD_Thai(sDay, sMonth, sYear, Optional dDelim As String = "") As String
Dim strD$, strM$, strY$
    strD = AddZero("" & sDay)
    strM = AddZero("" & sMonth)
    If sYear > 2500 Then
        strY = "" & sYear
    Else
        strY = CStr(sYear + 543)
    End If
    
    FormatYMD_Thai = Trim(strY) & dDelim & Trim(strM) & dDelim & Trim(strD)
End Function

Private Function FormatYMD_Eng(sDay%, sMonth%, sYear%, Optional dDelim As String = "") As String
Dim strD$, strM$, strY$
    strD = AddZero("" & sDay)
    strM = AddZero("" & sMonth)
    If sYear > 2500 Then
        strY = "" & (sYear - 543)
    Else
        strY = "" & sYear
    End If
    
    FormatYMD_Eng = Trim(strY) & dDelim & Trim(strM) & dDelim & Trim(strD)
End Function

Private Function FormatYMD_to_DMY(sDate1 As String, Optional dDelim1 As String = "/", Optional dDelim2 As String = "/") As String
Dim tmpArr
Dim tmpStr$
Dim d$, m$, y$
    If InStr(1, sDate1, dDelim1) = 0 Then
        FormatYMD_to_DMY = ""
        Exit Function
    End If
    
    tmpArr = Split(sDate1, dDelim1)
    
    d = AddZero(Trim("" & tmpArr(2)))
    m = AddZero(Trim("" & tmpArr(1)))
    y = Trim("" & tmpArr(0))
    
    FormatYMD_to_DMY = d & dDelim2 & m & dDelim2 & y
End Function

Private Function FormatDMY_to_YMD(sDate1 As String, dDelim1 As String, dDelim2 As String) As String
Dim tmpArr
Dim tmpStr$
Dim d$, m$, y$
    If InStr(1, sDate1, dDelim1) = 0 Then
        FormatDMY_to_YMD = ""
        Exit Function
    End If
    
    tmpArr = Split(sDate1, dDelim1)
    
    d = AddZero(Trim("" & tmpArr(0)))
    m = AddZero(Trim("" & tmpArr(1)))
    y = Trim("" & tmpArr(2))
    
    FormatDMY_to_YMD = y & dDelim2 & m & dDelim2 & d
End Function

