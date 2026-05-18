Attribute VB_Name = "modFileVersion"
Option Explicit

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
        If Trim(AppVersion) <> "" Then AppVerInt = (Int(Arr(0)) * 1000) + (Int(Arr(1)) * 100) + (Int(Arr(2)) * 10) + Int(Arr(3))
        Arr() = Split(NewVersion, ".")
        If Trim(NewVersion) <> "" Then NewVerInt = (Int(Arr(0)) * 1000) + (Int(Arr(1)) * 100) + (Int(Arr(2)) * 10) + Int(Arr(3))
        If NewVerInt > AppVerInt Then
            AppNewVersion = True
        Else
            AppNewVersion = False
        End If
End Function
