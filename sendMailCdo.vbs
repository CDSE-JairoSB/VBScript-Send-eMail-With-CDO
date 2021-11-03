Option Explicit
Dim Email
Set Email=CreateObject("CDO.Message")
'************************************** Introduce Data 
Email.Subject="Remplace_Subject_Name"
Email.From="Remplace_eMail_Transmitter"
Email.To="Remplace_eMail_Receiver"
Email.TextBody="Text_Into_Body_Mail"
Email.AddAttachment("Optional_Path_File(s)")
'************************************** Configuration Data for Server
Email.Configuration.Fields.Item("http://schemas.Microsoft.com/cdo/configuration/sendusing")=2
'SMTP Server
Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.office365.com" 'server of your email provider 
'SMTP Port
Email.Configuration.Fields.Item("http://schemas.Microsoft.com/cdo/configuration/smtpserverport")=587 'Other port is 25
'************************************** Update & Send eMail
Email.Configuration.Fields.Update
Email.Send

set MyEmail=nothing
