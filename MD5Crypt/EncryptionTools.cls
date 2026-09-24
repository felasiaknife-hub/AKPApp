VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncryptionTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Encrypt(ByVal PublicKey As String, ByRef Data As String) As String
    Encrypt = CryptData(PublicKey, Data, True)
End Function

Public Function Decrypt(ByVal PublicKey As String, ByRef Data As String) As String
    Decrypt = CryptData(PublicKey, Data, False)
End Function

Public Function MD5HashString(ByVal Data As String) As String
    MD5HashString = DigestStrToHexStr(Data)
End Function

Public Function MD5HashFile(ByVal Filename As String) As String
    MD5HashFile = DigestFileToHexStr(Filename)
End Function

Public Function URLEncode(ByRef Data As String)
    URLEncode = URLEncodeData(Data)
End Function

Public Function URLDecode(ByRef Data As String)
    URLDecode = URLEncodeData(Data)
End Function

Public Function URLEncodeBinaryData(ByRef Data As String)
    URLEncodeBinaryData = URLEncodeDataForFile(Data)
End Function

Public Function URLDecodeBinaryData(ByRef Data As String)
    URLDecodeBinaryData = URLDecodeDataForFile(Data)
End Function


