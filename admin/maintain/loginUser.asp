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
		If Request("pswrd") <> Trim(connEmail("pswrd")) Then
			badPassword = True
		Else
			Session("userID")=connEmail("userID")
			Session("firstName")=connEmail("firstName")
			Session("lastName")=connEmail("lastName")
			IF connEmail("noImages")=1 THEN Session("noImages")="True" ELSE Session("noImages")="False" END IF
			If Trim(connEmail("rights"))="admin" Then Session("validAdmin")=True End If
			If Trim(connEmail("rights"))="director" Then Session("validDirector")=True End If
			If Trim(connEmail("rights"))="inspector" Then Session("validInspector")=True End If
			If Trim(connEmail("rights"))="user" Then Session("validUser")=True End If
			If Trim(connEmail("rights"))="action" Then Session("validUser")=True End If
			If Trim(connEmail("rights"))="erosion" Then Session("validErosion")=True End If
            Session("seeScoring") = connEmail("seeScoring")
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='user' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validUser")=True END IF
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='inspector' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validInspector")=True END IF
			IF NOT(Session("validAdmin")) THEN
				IF Session("adminReturnTo")="" THEN Session("adminReturnTo") = "../../" END IF
			ELSE
				IF Session("adminReturnTo")="" THEN Session("adminReturnTo") = "../" END IF
			END IF
			connSWPPP.Close
			Set connSWPPP = Nothing
Response.Write(Session("adminReturnTo"))
			Response.Redirect(Session("adminReturnTo"))
		End If ' no matching password
	End If ' no matching email
End If ' Request.Form.Count>0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP INSPECTIONS :: Admin :: Login</title>
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
		<font color="#FF0000">Your email/password does not match our admin list.<br>
		Please resubmit.</font><br><br>
<% 	End If %>
<table bgcolor="#006699">
<tr><td colspan="2" align="center"><br><h1><font color="#FFffff">User Login</font></h1>
		</td></tr>
	<tr><td colspan="2" align="center"><br><h3><font color="#FFffff">Due to a security upgrade all passwords will need to be reset.<br/>  
	Click the lost password link below to generate a new password.<br/>After receiving the new password by email, <br/>click on the change your password link to set it to what you want.</font></h3>
		</td></tr>
	<tr><td colspan="2" bgcolor="#ff3333"><img src="../../images/dot.gif" width="1" height="1"
			border="0" alt=""></td></tr>
	<tr><td align="right"><br><font color="#FFFFFF">Email: </font></td>
		<td><br><input type="text" name="email" size="30" maxlength="50"></td></tr>
	<tr><td align="right"><font color="#FFFFFF">Password: </font></td>
		<td><input type="password" name="pswrd" size="8" maxlength="8"></td></tr>
	<tr><td colspan="2" align="center"><br><font color="#FFFFFF"><a href="lostPassword.asp" style="color:#FFFFFF; font-weight:600;">have you lost your password?</a>
			</font><br><br></td></tr>
	<tr><td colspan="2" align="center"><br><font color="#FFFFFF"><a href="changePassword.asp" style="color:#FFFFFF; font-weight:600;">change your password</a>
			</font><br><br></td></tr>
	<tr><td colspan="2" align="center"><br><input type="submit" value="Submit"><br><br></td></tr>
	</table></div>
	<script language="javascript"><!--
		document.forms[0].elements[0].focus();
	//--></script>
</form>
</body>
</html>