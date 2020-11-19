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
		Session("userID")=connUsers("userID")
		' call the function to generate random password
		strPassword = GenerateRandomPassword ()
		
		SQL0="UPDATE Users SET pswrd='"& strPassword &"' WHERE userID="& Session("userID") 
		SET RS0=connSWPPP.execute(SQL0)
	
		dim mailServer, sendEmail
		sendEmail = strPassword & vbCrLf
		firstName = Trim(connUsers("firstName"))
		lastName = Trim(connUsers("lastName"))
		connUsers.Close
		Set connUsers = Nothing

		Set Mailer = Server.CreateObject("Persits.MailSender")
		Mailer.FromName    = "Don Wims"
		Mailer.From        = "dwims@swppp.com"
		Mailer.Host        = "127.0.0.1"
		Mailer.Subject     = "password for swppp.com"
		Mailer.Body        = sendEmail
		Mailer.isHTML      = True
		
		Mailer.AddAddress Request("email"), firstName & " " & lastName
		On Error Resume Next
		Mailer.Send
		If Err <> 0 Then %>
			<FONT color="red">Mail send failure.- </FONT><%= Err.Description %><br>
<%		else %>
			<FONT color="red">An email was sent to your email address containing a new password. <a href="loginUser.asp">Return to login page.</a></FONT><br>
<%		end if

		connSWPPP.Close
		Set connSWPPP = Nothing
	End If ' no matching email
End If ' Request.Form.Count>0

function GenerateRandomPassword ()
dim intPWLength, intLoop, intCharType, strPwd
Const intMinPWLength = 8
Const intMaxPWLength = 8

' Generates a random number: 6, 7, 8, 9, or 10
' this number determines the length of the password. For instance, if
' the random number is 10 then, the password length will be 10
Randomize
intPWLength = int((intMaxPWLength - intMinPWLength + 1) * Rnd + intMinPWLength)
' now depending on the length of the password (dependent on the random
' number generated above), create random chracters between a-z, A-Z, or
' or 0-9 by using a for loop
for intLoop = 1 To intPWLength
' Generates a random number: 1, 2, or 3; where
' 1 gets a lowercase letter; 2 gets uppercase character, and
' 3 gets a number between 0 and 9
Randomize
intCharType = Int((3 * Rnd) + 1)

' now check if intCharType is 1, 2, or 3
select case intCharType
case 1
' get a lowercase letter a-z inclusive
Randomize
strPwd = strPwd & CHR(Int((25 * Rnd) + 97))
case 2
' get a uppercase letter A-Z inclusive
Randomize
strPwd = strPwd & CHR(Int((25 * Rnd) + 65))
case 3
' get a number between 0 and 9 inclusive
Randomize
strPwd = strPwd & CHR(Int((9 * Rnd) + 48))
end select
next

' return password
GenerateRandomPassword = strPwd
end function
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP :: Users :: login</title>
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
<tr><td colspan="2" align="center"><br><h1><font color="#FFffff">lost password</font></h1>
		<font color="#FFffff">Please enter your email address and<br>a new password will be emailed to you.
		</font><br><br></td></tr>
	<tr><td colspan="2" bgcolor="#ff3333"><img src="../../images/dot.gif" width="1" height="1"
			border="0" alt=""></td></tr>
	<tr><td align="right"><font color="#FFffff"><br>email: </td>
		<td><br><input type="text" name="email" size="30" maxlength="50">&nbsp;&nbsp;</td></tr>
	<tr><td colspan="2" align="center"><br><input type="submit" value="get password"><br><br></td></tr>
	</table></div>
	<script language="javascript"><!--
		document.forms[0].elements[0].focus();
	//--></script>
</form>

</body>
</html>
