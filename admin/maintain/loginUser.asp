<%
Session.Timeout=120 'session timeout
Session("userID")         = ""
Session("validAdmin")     = "False"
Session("validDirector")  = "False"
Session("validInspector") = "False"
Session("validUser")      = "False"
Session("validErosion")   = "False"
Session("validVSCR")      = "False"
Session("validLDSCR")     = "False"
Session("seeScoring")= "True"
If Request.Form.Count > 0 Then
	userSQLSELECT = "SELECT * FROM Users WHERE email = '" & Request("email") & "'"
	%> <!-- #INCLUDE FILE="../connSWPPP.asp" --> <%
	' Response.Write(userSQLSELECT & "<br>")
	Set connEmail = connSWPPP.execute(userSQLSELECT)
	If connEmail.EOF Then
		noMatch = True
	Else
		If Trim(Request("pswrd")) <> Trim(connEmail("pswrd")) Then
			badPassword = True
		Else If connEmail("active") == 0 Then
			inactiveUser = True
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
			If Trim(connEmail("rights"))="vscr" Then Session("validVSCR")=True End If
			If Trim(connEmail("rights"))="ldscr" Then Session("validLDSCR")=True End If
			Session("seeScoring") = connEmail("seeScoring")
			
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='user' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validUser")=True END IF

			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='inspector' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validInspector")=True END IF
			
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='vscr' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validVSCR")=True END IF
			
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='ldscr' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validLDSCR")=True END IF
			

			IF NOT(Session("validAdmin")) THEN
				IF Session("adminReturnTo")="" THEN Session("adminReturnTo") = "../../" END IF
			ELSE
				IF Session("adminReturnTo")="" THEN Session("adminReturnTo") = "../" END IF
			END IF
			
			connSWPPP.Close
			Set connSWPPP = Nothing
			'Response.Write(Session("adminReturnTo"))
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
<% If noMatch Then %>
		<font color="#FF0000">Your email cannot be found in our admin list. Please
		resubmit.</font><br><br>
<% End If
	If badPassword Then %>
		<font color="#FF0000">Your email/password does not match our admin list.<br>
		Please resubmit.</font><br><br>
<% End If
	If inactiveUser Then %>
		<font color="#FF0000">This account has been deleted. If you feel this is an error please contact us.</font><br><br>
<% End If %>
<table bgcolor="#006699">
<tr><td colspan="2" align="center"><br><h1><font color="#FFffff">user login</font></h1>
		</td></tr>
	<tr><td colspan="2" bgcolor="#ff3333"><img src="../../images/dot.gif" width="1" height="1"
			border="0" alt=""></td></tr>
	<tr><td align="right"><br><font color="#FFFFFF">email: </font></td>
		<td><br><input type="text" name="email" size="30" maxlength="50"></td></tr>
	<tr><td align="right"><font color="#FFFFFF">password: </font></td>
		<td><input type="password" name="pswrd" size="15" maxlength="15"></td></tr>
	<tr><td colspan="2" align="center"><br><font color="#FFFFFF"><a href="lostPassword.asp" style="color:#FFFFFF; font-weight:600;">have you lost your password?</a>
			</font><br><br></td></tr>
	<tr><td colspan="2" align="center"><br><font color="#FFFFFF"><a href="changePassword.asp" style="color:#FFFFFF; font-weight:600;">change your password</a>
			</font><br><br></td></tr>
	<tr><td colspan="2" align="center"><br><input type="submit" value="submit"><br><br></td></tr>
	</table></div>
	<script language="javascript"><!--
		document.forms[0].elements[0].focus();
	//--></script>
</form>
</body>
</html>