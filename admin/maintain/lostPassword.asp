<%
'smp 12/26/02

fromAddr = "info@swppp.com"
remoteHost= "127.0.0.1"
loginURL = "http://www.swppp.com/admin/maintain/loginUser.asp"

If Request.Form.Count > 0 Then
	SQLSELECT = "SELECT userID, pswrd, firstName, lastName" &_
		" FROM Users" & _
		" WHERE email = '" & Request("email") & "'"
	%> <!-- #INCLUDE FILE="../connSWPPP.asp" --> <%
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
<div align="center"><br><br>
<% 	If noMatch Then %>
		<font color="#FF0000">Your email cannot be found in our users list. Please
		resubmit.</font><br><br>
<% 	End If
	If badPassword Then %>
		<font color="#FF0000">Your email/password does not match our users list.<br>
		Please resubmit.</font><br><br>
<% 	End If %>

<table bgcolor="#006699">
<tr><td colspan="2" align="center"><br><h1><font color="#FFffff">Lost Password</font></h1>
		<font color="#FFffff">Please enter your email address and<br>your password will be emailed to you.
		</font><br><br></td></tr>
	<tr><td colspan="2" bgcolor="#ff3333"><img src="../../images/dot.gif" width="1" height="1"
			border="0" alt=""></td></tr>
	<tr><td align="right"><font color="#FFffff"><br>Email: </td>
		<td><br><input type="text" name="email" size="30" maxlength="50">&nbsp;&nbsp;</td></tr>
	<tr><td colspan="2" align="center"><br><input type="submit" value="Get Password"><br><br></td></tr>
	</table></div>
	<script language="javascript"><!--
		document.forms[0].elements[0].focus();
	//--></script>
</form>

</body>
</html>
