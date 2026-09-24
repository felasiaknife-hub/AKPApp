VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC4Crypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************
'*Class/Method Name:    RC4Crypt
'*Version:              1.0
'*File Name:            RC4Crypt.cls
'*Decription:           RC4 standard crypto-module for aggregator
'*                      - contains system RC4 calls
'*Author:               CPTjun
'*Company:              3rdFrontier (M) Sdn Bhd
'*Project:              3F Aggregator (Copyright (c)2001)
'*Date submitted:
'********************************************************************
'*MODIFICATION HISTORY
'* Name              Date        Version      Description
'*-------           ------      ---------    --------------
'* TEW              31/07       Beta 1.0    Integrated with Aggregator
'*                                          engine
'*
'*
'*
'********************************************************************/
'
'RC4 Encryption/Decryption Class
'------------------------------------
'
'Information concerning the RC4
'algorithm can be found at:
'http://www.rsasecurity.com/rsalabs/faq/3-6-3.html
'
'(c) 2000, Fredrik Qvarfort
'

Option Explicit

'For progress notifications
Event Progress(Percent As Long)

'Key-dependant data
Private m_Key As String
Private m_sBox(0 To 255) As Integer

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Sub DecryptByte(ByteArray() As Byte, Optional Key As String)

  'The same routine is used for encryption as well
  'decryption so why not reuse some code and make
  'this class smaller (that is it it wasn't for all
  'those damn comments ;))
  Call EncryptByte(ByteArray(), Key)

End Sub

Public Function EncryptString(Text As String, Optional Key As String, Optional FlagUnicode As Boolean = False) As String
    Dim ByteArray() As Byte
     
'    CurrentCulture = Localization.GetRegionalSettings
'    DefaultLCID = CurrentCulture
'    DefaultCalendarType = GetUserLCID(DefaultLCID, LOCALE_ICALENDARTYPE)
'    If CurrentCulture <> "1054" Then Localization.SetRegionalSettings "1054", CAL_THAI, "dd/MM/yyyy"
'
    If FlagUnicode = True Then
        'Convert the data into a byte array
        ByteArray() = StrConv(Text, vbFromUnicode, 1054)
        
        'Encrypt the byte array
        Call EncryptByte(ByteArray(), Key)
        
        'Convert the byte array back into a string
        EncryptString = StrConv(ByteArray(), vbFromUnicode, 1054)
    Else
        'Convert the data into a byte array
        ByteArray() = StrConv(Text, vbFromUnicode, 1054)
        
        'Encrypt the byte array
        Call EncryptByte(ByteArray(), Key)
    
        'Convert the byte array back into a string
        EncryptString = StrConv(ByteArray(), vbUnicode, 1054)
    End If
    
'    CurrentCulture = Localization.GetRegionalSettings
    
End Function

Public Function DecryptString(Text As String, Optional Key As String) As String

  Dim ByteArray() As Byte
 
  'Convert the data into a byte array
  ByteArray() = StrConv(Text, vbFromUnicode, 1054)
  
  If UBound(ByteArray) > 0 Then
    'Decrypt the byte array
    Call DecryptByte(ByteArray(), Key)
    
    'Convert the byte array back into a string
    DecryptString = StrConv(ByteArray(), vbUnicode, 1054)
  End If
End Function
Public Sub EncryptByte(ByteArray() As Byte, Optional Key As String)

  Dim i As Long
  Dim j As Long
  Dim Temp As Byte
  Dim Offset As Long
  Dim OrigLen As Long
  Dim CipherLen As Long
  Dim CurrPercent As Long
  Dim NextPercent As Long
  Dim sBox(0 To 255) As Integer
  
  'Set the new key (optional)
  If (Len(Key) > 0) Then Me.Key = Key
  
  'Create a local copy of the sboxes, this
  'is much more elegant than recreating
  'before encrypting/decrypting anything
  Call CopyMem(sBox(0), m_sBox(0), 512)
  
  'Get the size of the source array
  OrigLen = UBound(ByteArray) + 1
  
  CipherLen = OrigLen
  
  'Encrypt the data
  For Offset = 0 To (OrigLen - 1)
    i = (i + 1) Mod 256
    j = (j + sBox(i)) Mod 256
    Temp = sBox(i)
    sBox(i) = sBox(j)
    sBox(j) = Temp
    ByteArray(Offset) = ByteArray(Offset) Xor (sBox((sBox(i) + sBox(j)) Mod 256))
    
    'Update the progress if neccessary
    If (Offset >= NextPercent) Then
      CurrPercent = Int((Offset / CipherLen) * 100)
      NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
      RaiseEvent Progress(CurrPercent)
    End If
  Next

  'Make sure we return a 100% progress
  If (CurrPercent <> 100) Then RaiseEvent Progress(100)

End Sub

Public Property Let Key(New_Value As String)

  Dim a As Long
  Dim b As Long
  Dim Temp As Byte
  Dim Key() As Byte
  Dim KeyLen As Long
  
  'Do nothing if the key is buffered
  If (m_Key = New_Value) Then Exit Property
  
  'Set the new key
  m_Key = New_Value
  
  'Save the password in a byte array
  Key() = StrConv(m_Key, vbFromUnicode)
  KeyLen = Len(m_Key)
  
  'Initialize s-boxes
  For a = 0 To 255
    m_sBox(a) = a
  Next a
  For a = 0 To 255
    b = (b + m_sBox(a) + Key(a Mod KeyLen)) Mod 256
    Temp = m_sBox(a)
    m_sBox(a) = m_sBox(b)
    m_sBox(b) = Temp
  Next
  
End Property



