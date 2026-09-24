Attribute VB_Name = "Module1"
Option Explicit

Public Const ENC_KEY As String = "12638546654"

Public Sub main()
    'TodayDate = FormatYMD_Thai(Day(Date), Month(Date), Year(Date), "/")
    
    Set DBManager = New AWSDBManager
    DBConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Krisana\Projects\Common\UserSecurity\MSAccess\ModSecurity.mdb;Persist Security Info=False"
    DBManager.ConnectionString = DBConString
    frMain .Show
End Sub
