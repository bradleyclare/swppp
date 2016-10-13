<%
Session("userID")= ""
Session("validAdmin")= "False"
Session("validDirector")= "False"
Session("validInspector")= "False"
Session("validUser")= "False"
Session("validErosion")= "False"
If Request.Form.Count > 0 Then
	userSQLSELECT = "SELECT userID, pswrd, rights, firstName, lastName, noImages" &_
		" FROM Users" & _
		" WHERE email = '" & Request("email") & "'"
%>
<!-- #include virtual="admin/connSWPPP.asp" -->
<%
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
			IF connEmail("noImages")=1 THEN Session("noImages")="True" ELSE Session("noImages")="False" End If
			If Trim(connEmail("rights"))="admin" Then Session("validAdmin")=True End If
			If Trim(connEmail("rights"))="director" Then Session("validDirector")=True End If
			If Trim(connEmail("rights"))="inspector" Then Session("validInspector")=True End If
			If Trim(connEmail("rights"))="user" Then Session("validUser")=True End If
			If Trim(connEmail("rights"))="action" Then Session("validUser")=True End If
			If Trim(connEmail("rights"))="erosion" Then Session("validErosion")=True End If
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='user' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validUser")=True End If
			SQL0="SELECT COUNT(*) FROM ProjectsUsers WHERE rights='inspector' AND userID="& Session("userID")
			SET RS0=connSWPPP.execute(SQL0)
			IF RS0(0)>0 THEN Session("validInspector")=True End If
			IF NOT(Session("validAdmin")) THEN
				IF Session("adminReturnTo")="" THEN Session("adminReturnTo") = "../../" End If
			ELSE
				IF Session("adminReturnTo")="" THEN Session("adminReturnTo") = "../" End If
			End If
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
    <form action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
        <div class="login-div">
            <% 	If noMatch Then %>
            <h3>Your email cannot be found in our admin list. Please resubmit.</h3>
            <% 	End If
			If badPassword Then %>
            <h3>Your email/password does not match our admin list. Please resubmit.</h3>
            <% 	End If %>
            <h1>User Login</h1>
			<div class="row">
				<div class="four columns alpha"><h4>Email:</h4></div>
				<div class="eight columns omega"><input type="text" name="email"></div>
			</div>
			<div class="row">
				<div class="four columns alpha"><h4>Password:</h4></div>
				<div class="eight columns omega"><input type="password" name="pswrd"></div>
            </div>
            <a href="lostPassword.asp"><h3>Forgot Password?</h3></a>
			<div class="row">
				<button type="submit">Submit</button>
			</div>
        </div>
        <script language="javascript"><!--
    document.forms[0].elements[0].focus();
    //--></script>
    </form>
</body>
</html>

