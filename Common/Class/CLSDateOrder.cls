VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSDateOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DateOrderEnum
   doDefault 'Your locale setting
   doMDY     'Month-Day-Year (U.S.)
   doDMY     'Day-Month-Year (EU, S.A.)
   doYMD     'Year-Month-Day (Japan)
End Enum

Private Const LOCALE_SSHORTDATE As Long = &H1F
Private Const LOCALE_STHOUSAND As Long = &HF
Private Const LOCALE_SDECIMAL  As Long = &HE

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Function GetThousandsSep() As String
   GetThousandsSep = pfGLI(GetUserDefaultLCID(), LOCALE_STHOUSAND)
End Function

Public Function GetDecimalSep() As String
   GetDecimalSep = pfGLI(GetUserDefaultLCID(), LOCALE_SDECIMAL)
End Function

'Purpose: Assume a date string with English separator "1/4/2006"
'Returns: Correct Date Variable
Public Function ResolveDate(ByVal sDate As String) As Date
   Dim sArray() As String
   If InStr(sDate, "/") Then 'Potentially a date string
      sArray = Split(sDate, "/")
      Debug.Print "GetUserDefaultLCID", GetUserDefaultLCID
      Debug.Print "GetSystemDefaultLCID", GetSystemDefaultLCID
      If UBound(sArray) = 2 Then 'We have 3 parts
         Select Case ShortDateOrder2
            Case doMDY '
               ResolveDate = DateSerial(sArray(2), sArray(0), sArray(1))
            Case doDMY
               ResolveDate = DateSerial(sArray(2), sArray(1), sArray(0))
            Case doYMD
               ResolveDate = DateSerial(sArray(0), sArray(1), sArray(2))
         End Select
      End If
   End If
End Function

'Purpose: Assume a number string with English separators "123,456.78"
'Returns: Correct Double Variable
Public Function ResolveNumber(ByVal sNum As String) As Double
   Dim sTS As String
   Dim sDS As String
   sTS = GetThousandsSep
   sDS = GetDecimalSep

   If (sTS = ",") And (sDS = ".") Then 'English
      'format is OK
   Else
      Dim i As Long
      Dim sMid As String
      For i = 1 To Len(sNum)
         Select Case Mid(sNum, i, 1)
            Case ","
               Mid(sNum, i, 1) = sTS
            Case "."
               Mid(sNum, i, 1) = sDS
         End Select
      Next
   End If

   ResolveNumber = CDbl(sNum)

End Function

Public Function ShortDateOrder2() As DateOrderEnum
   'Get ShortDateOrder the hard way
   Dim sShort           As String
   Dim qOn              As Boolean
   Dim i                As Integer
   Dim sChar            As String

   On Error Resume Next

   'Get the Short Date format
   sShort = pfGLI(GetUserDefaultLCID(), LOCALE_SSHORTDATE)

   For i = 1 To Len(sShort)
      sChar = Mid(sShort, i, 1)
      'Ignore items in single quotes (if any)
      If sChar = "'" Then
         qOn = Not qOn
      Else
         If Not qOn Then
            Select Case sChar
               Case "d"
                  ShortDateOrder2 = doDMY
                  Exit Function
               Case "m"
                  ShortDateOrder2 = doMDY
                  Exit Function
               Case "y"
                  ShortDateOrder2 = doYMD
                  Exit Function
            End Select
         End If
      End If
   Next
End Function

Private Function pfGLI(ByVal m_LocaleLCID As Long, ByVal reqInfo As Long) As String
   Dim Buffer As String * 255
   GetLocaleInfoA m_LocaleLCID, reqInfo, Buffer, 255
   pfGLI = StripNull(Buffer)
End Function

Public Function StripNull(ByVal StrIn As String) As String
   Dim nul              As Long
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         StripNull = Left$(StrIn, nul - 1)
      Case 1
         StripNull = ""
      Case 0
         StripNull = Trim$(StrIn)
   End Select
End Function
