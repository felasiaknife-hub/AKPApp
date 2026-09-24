Attribute VB_Name = "modEncoding"
Option Explicit

Public Function URLEncodeData(ByRef Data As String) As String
Dim strData     As String
Dim lngIndex    As Integer
Dim strHex      As String
    For lngIndex = 1 To Len(Data)
        Select Case Asc(Mid$(Data, lngIndex, 1))
            Case 48 To 57, 65 To 90, 97 To 122
                strData = strData & Mid$(Data, lngIndex, 1)
            Case 32
                strData = strData & "%" & Hex$(32)
            Case Else
                If Mid(Data, lngIndex, 1) = ":" Then
                    Debug.Print 1
                End If
                strHex = Hex$(Asc(Mid$(Data, lngIndex, 1)))
                If Len(strHex) = 1 Then
                    strHex = "0" & strHex
                End If
                strData = strData & "%" & strHex
        End Select
    Next lngIndex
    URLEncodeData = strData
End Function

Public Function URLDecodeData(ByRef Data As String) As String
Dim strData     As String
Dim lngIndex    As Integer
    lngIndex = 1
    Do Until lngIndex - 1 = Len(Data)
        Select Case Mid(Data, lngIndex, 1)
            Case "%"
                strData = strData & Chr$(Val("&H" & Mid(Data, lngIndex + 1, 2)))
                lngIndex = lngIndex + 3
            Case Else
                strData = strData & Mid(Data, lngIndex, 1)
                lngIndex = lngIndex + 1
        End Select
    Loop
    URLDecodeData = strData
End Function

Public Function URLEncodeDataForFile(ByRef Data As String) As String
Dim bytData()   As Byte
Dim strData     As String
Dim lngIndex    As Integer
Dim strHex      As String
    lngIndex = 0
    bytData = Data
    For lngIndex = 0 To UBound(bytData)
        strHex = Hex$(bytData(lngIndex))
        If Len(strHex) = 1 Then
            strHex = "0" & strHex
        End If
        strData = strData & strHex
    Next lngIndex
    URLEncodeDataForFile = strData
End Function

Public Function URLDecodeDataForFile(ByRef Data As String) As String
Dim bytData()   As Byte
Dim lngIndex    As Integer
    ReDim bytData((Len(Data) / 2) - 1)
    For lngIndex = 0 To UBound(bytData)
        bytData(lngIndex) = CLng("&H" & Mid(Data, (lngIndex * 2) + 1, 2))
    Next lngIndex
    URLDecodeDataForFile = bytData
End Function
