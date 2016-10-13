<%

Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName   = "SWPPP Inspections"
Mailer.FromAddress= "noreply@swppp.com"
Mailer.RemoteHost = "127.0.0.1"
Mailer.AddRecipient "Brad Leishman", "brad.leishman@outlook.com"
Mailer.Subject    = "SMTP works!"
Mailer.BodyText   = "Dear Brad" & VbCrLf & "The SMTP is sending!"
if Mailer.SendMail then
  Response.Write "Mail sent..."
else
  Response.Write "Mail send failure. Error was " & Mailer.Response
end if
%>