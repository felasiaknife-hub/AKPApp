Attribute VB_Name = "modEMail"
Option Explicit

Dim MailMsg, MailConf, MCFields
Const cdoBasic = 1

Public Sub sendEMail(EMailIDs As String, LocalIP As String)
Call AddSeverMsg("Sending startup mail...")

Set MailMsg = CreateObject("CDO.Message")
Set MailConf = CreateObject("CDO.Configuration")
Set MCFields = MailConf.Fields

'Configure SMTP Mail
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  '  SMTP
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.net4india.com" 'put SMTP Server here
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "********" 'put Pwd here
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "clients@sofnetsystems.net" 'put User Name here
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
MCFields.Update

Set MailMsg.Configuration = MailConf
'MailMsg.HTMLBody = txtMessage.Text
MailMsg.TextBody = LocalIP & vbCrLf & vbCrLf & vbCrLf & "---------" & vbCrLf _
    & "This is an automated mail. Please do not reply to this mail." & vbCrLf & vbCrLf _
    & "This e-mail" _
    & "correspondence is confidential and intended for the named recipient(s) only." _
    & "If you are not the named recipient and received this correspondence in error," _
    & "you must not copy, distribute or take any action in reliance on it."
MailMsg.Subject = "DGAC Server IP"
MailMsg.To = EMailIDs ' "ABCD <abcd@abcd.com>"
MailMsg.From = "DGAC Server <clients@sofnetsystems.net>"
'MailMsg.AddAttachment "C:\Documents and Settings\Sofnet\Desktop\Test.txt"

On Error GoTo err
MailMsg.Send
Call AddSeverMsg("The startup mail has been sent successfully")
Exit Sub

err:
Call AddSeverMsg("ERROR: The startup mail could not be sent successfully")
End Sub




