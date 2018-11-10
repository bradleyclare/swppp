<%

Set Mail = Server.CreateObject("Persits.MailSender")


Mail.FromName = "SWPPP Inspections"
Mail.From = "noreply@swppp.com"
Mail.Host = "127.0.0.1"

Mail.AddAddress "bradleyclare@gmail.com", "Brad Leishman"
'Mail.AddCC
'Mail.AddBCC
Mail.Subject    = "SMTP works!"
Mail.Body   = "Dear Jeremy" & VbCrLf & "The SMTP is sending!"
Mail.isHTML      = True

On Error Resume Next
Mail.Send
If Err <> 0 Then
  Response.Write "An error occurred: " & Err.Description
Else
  Response.Write "Mail sent..."
End If

%>