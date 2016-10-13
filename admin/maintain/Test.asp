<%

On Error Resume Next

Session("userID")= ""
Session("validAdmin")= "False"
Session("validDirector")= "False"
Session("validInspector")= "False"
Session("validUser")= "False"

Response.Write("Here")
Response.End

If Request.Form.Count > 0 Then
	userSQLSELECT = "SELECT userID, pswrd, rights, firstName, lastName, noImages" &_
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
 