Attribute VB_Name = "modMsg"
Option Explicit

Public Function GetMsg(Optional ByVal StrMsg As String = "", Optional ByVal MsgBoxStyle As VbMsgBoxStyle = vbInformation) As String
Dim Msg As String
    If StrMsg = "" Then
        Msg = "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ"
    Else
        Msg = StrMsg
    End If
    MsgBox Err.Description & vbCrLf & Msg, MsgBoxStyle
End Function
