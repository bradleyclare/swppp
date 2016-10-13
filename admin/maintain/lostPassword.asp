<%
'smp 12/26/02

fromAddr = "info@swppp.com"
remoteHost= "127.0.0.1"
loginURL = "http://www.swppp.com/admin/maintain/loginUser.asp"

If Request.Form.Count > 0 Then
	SQLSELECT = "SELECT userID, pswrd, firstName, lastName" &_
		" FROM Users" & _
		" WHERE email = '" & Request("email") & "'"
	%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
'	Response.Write(SQLSELECT & "<br>")
	Set connUsers = connSWPPP.execute(SQLSELECT)

	If connUsers.EOF Then
		noMatch = True
	Else
		dim mailServer, sendEmail
		sendEmail = connUsers("pswrd") & vbCrLf
		firstName = Trim(connUsers("firstName"))
		lastName = Trim(connUsers("lastName"))
		connUsers.Close
		Set connUsers = Nothing

		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.FromName   = "SWPPP"
		Mailer.FromAddress= fromAddr
		Mailer.RemoteHost = remoteHost
		Mailer.AddRecipient firstName & " " & lastName, Request("email")
		Mailer.Subject    = "password for swppp.com"
		Mailer.BodyText   = sendEmail

		if not Mailer.SendMail then
		  Response.Write "Mail send failure. Error was " & Mailer.Response
		end if

'Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'Mailer.FromName   = "SWPPP"
'Mailer.FromAddress= "info@swppp.com"
'Mailer.RemoteHost = "127.0.0.1"
'Mailer.AddRecipient firstName & " " & lastName, Request("email")
'Mailer.AddRecipient "test", Request(email)
'Mailer.AddCC "jeremy zuther", "jzuther@gmail.com"
'Mailer.Subject    = "password for swppp.com"
'Mailer.BodyText   = sendEmail
'if Mailer.SendMail then
'  Response.Write "Mail sent..."
'else
'  Response.Write "Mail send failure. Error was " & Mailer.Response
'end if





		connSWPPP.Close
		Set connSWPPP = Nothing
		Response.Redirect("loginUser.asp")
	End If ' no matching email
End If ' Request.Form.Count>0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP :: Users :: Login</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>

<form action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
		<div class="login-div">
            <% 	If noMatch Then %>
            <h3>Your email cannot be found in our admin list. Please resubmit.</h3>
            <% 	End If
			If badPassword Then %>
            <h3>Your email/password does not match our admin list. Please resubmit.</h3>
            <% 	End If %>
            <h1>Forgot Password</h1>
			<h3>Please enter your email address and your password will be emailed to you.</h3>
			<div class="row">
				<div class="four columns alpha"><h4>Email:</h4></div>
				<div class="eight columns omega"><input type="text" name="email"></div>
            </div>
			<div class="row">
				<button type="submit">Get Password</button>
			</div>
        </div>
	<script language="javascript"><!--
		document.forms[0].elements[0].focus();
	//--></script>
</form>

</body>
</html>
