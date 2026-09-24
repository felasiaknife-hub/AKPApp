Attribute VB_Name = "modStreamFiles"
Option Explicit

Public Sub StreamLog(ByVal LogFile As String, ByVal MainStr As String, Optional ByVal OptionStr As String = "")
Dim ObjStream As New FileSystemObject
Dim CreateLog As TextStream
Dim AppPath As String
Dim FilePath As String
    AppPath = App.Path & "\"
    FilePath = AppPath + LogFile
    If ObjStream.FileExists(FilePath) = False Then
        Set CreateLog = ObjStream.CreateTextFile(FilePath, False)
    Else
        Set CreateLog = ObjStream.OpenTextFile(FilePath, ForAppending, False, TristateUseDefault)
    End If
    CreateLog.WriteLine (MainStr & OptionStr)
End Sub
