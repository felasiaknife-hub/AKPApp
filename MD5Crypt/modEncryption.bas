Attribute VB_Name = "modEncryption"
Option Explicit

Private Const LOWER_BYTE_VALUE      As Long = 48
Private Const UPPER_BYTE_VALUE      As Long = 150

Public Function CryptData(ByVal Key As String, ByRef Data As String, ByVal pblnEncrypt As Boolean) As String
Dim lngIndex        As Integer
Dim lngKeyLen       As Long
Dim lngKeyPos       As Long
Dim bytData()       As Byte
Dim bytKey()        As Byte
Dim bytOffset()     As Byte
Dim lngCode         As Long
Dim lngAvg          As Long
    bytData = Data
    lngKeyLen = Len(Key)
    ReDim bytKey(lngKeyLen - 1)
    ReDim bytOffset(lngKeyLen - 1)
    For lngIndex = 0 To lngKeyLen - 1
        bytKey(lngIndex) = Asc(Mid$(Key, lngIndex + 1, 1))
        bytOffset(lngIndex) = bytKey(lngIndex) / 2
        lngAvg = lngAvg + bytKey(lngIndex)
    Next lngIndex
    lngAvg = lngAvg / lngKeyLen
    For lngIndex = 0 To lngKeyLen - 1
        bytKey(lngIndex) = (bytKey(lngIndex) + lngAvg) / 2
    Next lngIndex
    For lngIndex = 0 To UBound(bytData)
        lngKeyPos = (lngIndex Mod lngKeyLen)
        lngCode = bytData(lngIndex)
        If pblnEncrypt Then
            lngCode = lngCode Xor bytKey(lngKeyPos)
        End If
        If ((lngCode >= LOWER_BYTE_VALUE) And (lngCode <= UPPER_BYTE_VALUE)) Then
            If pblnEncrypt Then
                lngCode = lngCode + bytOffset(lngKeyPos)
            Else
                lngCode = lngCode - bytOffset(lngKeyPos)
            End If
            If lngCode > UPPER_BYTE_VALUE Then
                lngCode = lngCode - UPPER_BYTE_VALUE + LOWER_BYTE_VALUE - 1
            ElseIf lngCode < LOWER_BYTE_VALUE Then
                lngCode = lngCode + UPPER_BYTE_VALUE - LOWER_BYTE_VALUE + 1
            End If
        End If
        If Not pblnEncrypt Then
            lngCode = lngCode Xor bytKey(lngKeyPos)
        End If
        bytData(lngIndex) = lngCode
    Next lngIndex
    CryptData = bytData
End Function
