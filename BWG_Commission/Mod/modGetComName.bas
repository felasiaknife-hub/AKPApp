Attribute VB_Name = "modGetComName"
Option Explicit

Private Declare Function apiGetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function fOSMachineName() As String
'Returns the computername
Dim lngLen As Long, lngX As Long
Dim strCompName As String
    lngLen = 16
    strCompName = String$(lngLen, 0)
    lngX = apiGetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        fOSMachineName = Left$(strCompName, lngLen)
    Else
        fOSMachineName = ""
    End If
End Function

