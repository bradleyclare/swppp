<%
Session("userID")= ""
Session("validAdmin")= "False"
Session("validDirector")= "False"
Session("validInspector")= "False"
Session("validUser")= "False"
Session("validErosion")= "False"
Session("seeScoring")= "True"
If Request.Form.Count > 0 Then
	userSQLSELECT = "SELECT userID, pswrd, rights, firstName, lastName, noImages, seeScoring" &_
		" FROM Users" & _
		" WHERE email = '" & Request("email") & "'"
	%> <!-- #INCLUDE FILE="../connSWPPP.asp" --> <%
	' Response.Write(userSQLSELECT & "<br>")
	Set connEmail = connSWPPP.execute(userSQLSELECT)
	If connEmail.EOF Then
		noMatch = True
	Else
		If Request("current_pswrd") <> Trim(connEmail("pswrd")) Then
			badPassword = True
		Else
			If Request("new_pswrd") <> Request("new_pswrd2") Then
				newNoMatch = True
			ELSE
				Session("userID")=connEmail("userID")
				SQL0="UPDATE Users SET pswrd='"& Request("new_pswrd") &"' WHERE userID="& Session("userID") 
				SET RS0=connSWPPP.execute(SQL0)
				connSWPPP.Close
				Set connSWPPP = Nothing
				success_flag = True
			End IF 'new passwords do not match
		End If ' no matching password
	End If ' no matching email
End If ' Request.Form.Count>0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP INSPECTIONS :: Admin :: Change Password</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<body>
<span id="siteseal"><script async type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=WAZqkjKfwncrXQiy57BnnkkIp0xnpa50j7Om4owvXUaaZQu6tQU4wBV9R1iL"></script></span>
<form action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
<div align="center"><br><br>
<% 	If noMatch Then %>
		<font color="#FF0000">Your email cannot be found in our admin list. Please
		resubmit.</font><br><br>
<% 	End If
	If badPassword Then %>
		<font color="#FF0000">Your email and current password does not match.<br>
		Please resubmit.</font><br><br>
<% 	End If
	If newNoMatch Then %>
		<font color="#FF0000">Your new passwords do not match.<br>
		Please resubmit.</font><br><br>
<% 	End If
	If success_flag Then %>
		<font color="#FF0000">Your password has been changed. <a href='../../default.asp'>Click here to return to the home page.</a></font><br><br>
<% 	End If %>
<table bgcolor="#006699">
<tr><td colspan="2" align="center"><br><h1><font color="#FFffff">Change Password</font></h1>
		</td></tr>
	<tr><td colspan="2" bgcolor="#ff3333"><img src="../../images/dot.gif" width="1" height="1"
			border="0" alt=""></td></tr>
	<tr><td align="right"><br><font color="#FFFFFF">Email: </font></td>
		<td><br><input type="text" name="email" size="30" maxlength="50"></td></tr>
	<tr><td align="right"><font color="#FFFFFF">Current Password: </font></td>
		<td><input type="password" name="current_pswrd" size="15" maxlength="15"></td></tr>
	<tr><td align="right"><font color="#FFFFFF">New Password: </font></td>
		<td><input type="password" name="new_pswrd" size="15" maxlength="15"></td></tr>
	<tr><td align="right"><font color="#FFFFFF">New Password Repeat: </font></td>
		<td><input type="password" name="new_pswrd2" size="15" maxlength="15"></td></tr>
	<tr><td colspan="2" align="center"><br><input type="submit" value="Submit"><br><br></td></tr>
	</table></div>
	<script language="javascript"><!--
		document.forms[0].elements[0].focus();
	//--></script>
</form>
</body>
</html>