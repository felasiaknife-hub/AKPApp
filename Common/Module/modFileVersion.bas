Attribute VB_Name = "modFileVersion"
Option Explicit

Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
  "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize _
  As Long) As Long
  
  Public Function GetSysDir() As String

'*******************************************************
'<DESC>     Retrieves System Dir Path</DESC>
'<RETURN>   System Directory Path</RETURN>
'<ACCESS>   Public
'<USAGE>    strSysPath = GetSysDir()</USAGE>
'*******************************************************
    
    Dim strBuf As String * 256
    Dim lngReturn_len As Long
    
    lngReturn_len = GetSystemDirectory(strBuf, Len(strBuf))
    GetSysDir = Trim$(Left$(strBuf, lngReturn_len))
    
End Function

Public Function AppNewVersion(ByVal CurPath As String, ByVal SourcePath As String, ByVal FileName As String) As Boolean
Dim CurApp As Scripting.FileSystemObject
Dim AppVersion As String, AppVerInt As Integer
Dim NewVersion As String, NewVerInt As Integer

        Set CurApp = New Scripting.FileSystemObject
        AppVersion = CurApp.GetFileVersion(CurPath & "\" & FileName)
        Set CurApp = Nothing
        Set CurApp = New Scripting.FileSystemObject
        NewVersion = CurApp.GetFileVersion(SourcePath & "\" & FileName)
        Set CurApp = Nothing
        
        Dim Arr() As String
        Arr() = Split(AppVersion, ".")
        If Trim(AppVersion) <> "" Then AppVerInt = (Int(Arr(0)) * 10000) + (Int(Arr(1)) * 1000) + (Int(Arr(2)) * 100) + Int(Arr(3))
        Arr() = Split(NewVersion, ".")
        If Trim(NewVersion) <> "" Then NewVerInt = (Int(Arr(0)) * 10000) + (Int(Arr(1)) * 1000) + (Int(Arr(2)) * 100) + Int(Arr(3))
        If NewVerInt > AppVerInt Then
            AppNewVersion = True
        Else
            AppNewVersion = False
        End If
End Function




