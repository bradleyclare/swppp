<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

%> <!-- #include file="../connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	highestRights="user"
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	
	Function titleCase(strValue)
		strValue=Ucase(Left(strValue,1)) & Lcase(Mid(strValue,2,Len(strValue)-1))
		titleCase=Replace(strValue,"'","''")
	end function
	
	' is user already in DB?
	userSQLSELECT = "SELECT * FROM Users WHERE email = '" & Request("email") & "'"
	Set connUser = connSWPPP.Execute(userSQLSELECT)
	
	If Not connUser.EOF Then
		userExists = True 
		connUser.Close
		Set connUser = Nothing
		connSWPPP.Close
		Set connSWPPP = Nothing
	Else
		IF IsNull(Request("qualifications")) THEN Request("qualifications")="" END IF
		trimmedQualifications=REPLACE(Request("qualifications"),"'","#@#")
		userSQLINSERT = "INSERT INTO Users (firstName, lastName, email" & _
			", pswrd, dateEntered, signature, noImages, rights, qualifications" & _
			") VALUES (" & _
			"'" & titleCase(Request("firstName")) & "'" & _
			", '" & titleCase(Request("lastName")) & "'" & _
			", '" & strQuoteReplace(Request("email")) & "'" & _
			", '" & strQuoteReplace(Request("pswrd")) & "'" & _
			", '" & date & "'" & _
			", '" & Request("signature") & "'" & _
			", '" & Request("noImages") & "'" & _
			", 'user'" & _
			", '" & trimmedQualifications & "'" & _
			")"
			
'Response.Write(userSQLINSERT & "<br>")
		connSWPPP.Execute(userSQLINSERT)
		
		maxSQLSELECT = "SELECT MAX(userID) FROM Users"
		Set connUser = connSWPPP.Execute(maxSQLSELECT)
		userIdentity = connUser(0)
		connUser.Close
		Set connUser = Nothing
		highestRights="user"
		If Request("admin")="on" then highestRights="admin" End If

' ----------------------- Inspector, Director, User in Companies User  ----------------------- 

		For Each Item in Request.Form
'Response.Write("Item=" & Item & " Request(Item)=" & Request(Item) & "<br>")
			Select Case Left(Item,3)
			Case "ins"
				rights="inspector"
				IF highestRights <>"admin" AND highestRights<>"director" THEN highestRights="inspector"
			Case "dir"
				' check that project doesn't already have a director
				SQLSELECT = "SELECT * FROM ProjectsUsers" &_
					" WHERE projectID=" & Request(Item) &_
					" AND rights='director'"
				Set connDir=connSWPPP.execute(SQLSELECT)
				If not connDir.eof then
					projectHasDir=true
				else
					rights="director"
					IF highestRights <> "admin" THEN highestRights="director"
				end if
				connDir.close
				Set connDir=Nothing
			Case "use"
				rights="user"
			End Select
			
			If rights="inspector" or rights="director" or rights="user" then
				SQLINSERT = "INSERT INTO ProjectsUsers (userID, projectID, rights) VALUES (" &_
					userIdentity &_
					", " & Request(Item) &_
					", '" & rights & "'" &_
					")"
Response.Write(SQLINSERT & "<br>")
				connSWPPP.Execute(SQLINSERT)
			end if 'item=inspector, director or user
			rights=""
		Next
			
		SQL0="UPDATE Users SET rights='"& highestRights &"'" &_
		" WHERE userID='"& userID &"'"
		connSWPPP.execute(SQL0)
		connSWPPP.Close
		Set connSWPPP = Nothing

		Response.Redirect("viewUsersAdmin.asp")
	End If ' is user already in DB?
End If
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Add User</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
	<script language="JavaScript" src="../js/validUsers.js"></script>
	<script language="JavaScript" src="../js/validUsers1.2.js"></script>
</head>

<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><h1>Add User</h1></td></tr></table>
<% If userExists Then %>
	<table width="100%" border="0">
		<tr><td align="center"><font color="red">User email address already exists in database. 
				Go to <a href="viewUsersAdmin.asp">View Users</a> to edit.</font></td></tr></table>
<% Else ' Not userExists %>
<table width="100%" border="0">
	<form action="<% = Request.ServerVariables("script_name") %>" method="post" 
		onSubmit="return isReady(this)";>
		<tr><td width="35%" align="right">First Name:</td>
			<td width="65%"><input type="text" name="firstName" size="20" maxlength="20"></td></tr>
		<tr><td align="right">Last Name:</td>
			<td><input type="text" name="lastName" size="20" maxlength="20"></td></tr>
		<tr><td align="right">Email:</td>
			<td><input type="text" name="email" size="30" maxlength="50"></td></tr>
		<tr><td align="right">Password:</td>
			<td><input type="password" name="pswrd" size="8" maxlength="8"></td></tr>
		<tr><td align="right">View Images:</td>
			<td><input type="radio" name="noImages" value="0"<% IF noImages=0 THEN %> checked<% END IF%>>Yes
				<input type="radio" name="noImages" value="1"<% IF noImages=1 THEN %> checked<% END IF%>>No</td></tr>
<% If Session("validAdmin") then '-- only admin may set inspectors signature files and qualifications %>
		<tr><td align="right">Signature File:</td>
			<td><select name="signature">
<% ' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder("d:\vol\swpppinspections.com\www\htdocs\images\signatures\")
Set gifDirectory = objFolder.Files

For Each gifFile In gifDirectory
	shortenedName = gifFile.Name %>
					<option value="<% = shortenedName %>"
						<% if shortenedName = "dot.gif" Then %> selected<% End If %>><% = shortenedName %>
					</option>
<% Next
Set objFolder = Nothing
Set gifDirectory = Nothing %>
				</select>&nbsp;&nbsp;<input type="button" value="Upload Signature File" 
				onClick="location='upSigAddUser.asp'; return false";></td></tr>
		<tr><td align="right" valign=top>Qualifications:</td>
			<td><TEXTAREA cols="50" rows="3" name="qualifications"></TEXTAREA></td></tr>
<% END IF '-- Valid ADMIN. %>
<!--- ----------------------------------------- Rights --------------------------------------- --->

		<tr><td align="left"><br><font size="+1">Rights</font><br><br></td></tr>

<% If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>		

		<tr><td align="right">Admin:</td>
			<td><input type="checkbox" name="admin"></td></tr>
		
<!--- ----------------------------------------- Inspector --------------------------------------- --->

		<tr><td align="right" valign="top">Inspector:</td>
			<td><table cellpadding="0" cellspacing="0" border="0">
<% compCount=0
SQLSELECT = "SELECT projectID, projectName" & _
	" FROM Projects" &_
	" ORDER BY projectName"
Set connComp = connSWPPP.Execute(SQLSELECT)

Do While Not connComp.EOF
	If Trim(connComp("projectName"))<>"" then
		compCount=compCount+1 %>
			<tr><td><%= Trim(connComp("projectName")) %></td>
				<td><input type="checkbox" name="ins<%= compCount %>" 
					value="<%= connComp("projectID") %>"></td></tr>
<%	end if 'blank company name
	connComp.MoveNext
Loop %>
				</table></td></tr>

<!--- ----------------------------------------- Director --------------------------------------- --->

		<tr><td colspan="2"><hr size="1" noshade></td></tr>
		<tr><td align="right" valign="top">Director:</td>
			<td><table cellpadding="0" cellspacing="0" border="0">
<%connComp.movefirst
compCount=0
Do While Not connComp.EOF
	If Trim(connComp("projectName"))<>"" then
		compCount=compCount+1 %>
					<tr><td><%= Trim(connComp("projectName")) %></td>
						<td><input type="checkbox" name="dir<%= compCount %>" 
							value="<%= connComp("projectID") %>"></td></tr>
<%	end if 'blank company name
	connComp.MoveNext
Loop %>
				</table></td></tr>

<!--- ----------------------------------------- User --------------------------------------- --->
<% 	connComp.movefirst
else 'select the companies for which this user is a validDirector
	SQLSELECT = "SELECT pu.projectID, p.projectName" & _
		" FROM ProjectsUsers as pu, Projects as p" &_
		" WHERE userID=" & Session("userID") &_
		" AND rights='director'" &_
		" AND pu.projectID=p.projectID" &_
		" ORDER BY p.projectName"
	Set connComp = connSWPPP.Execute(SQLSELECT)
end if 'Session("validAdmin") %>

		<tr><td colspan="2"><hr size="1" noshade></td></tr>
		<tr><td align="right" valign="top">User:</td>
			<td><table cellpadding="0" cellspacing="0" border="0">
<%
compCount=0
Do While Not connComp.EOF
	If Trim(connComp("projectName"))<>"" then
		compCount=compCount+1 %>
					<tr><td><%= Trim(connComp("projectName")) %></td>
						<td><input type="checkbox" name="use<%= compCount %>" 
							value="<%= connComp("projectID") %>"></td></tr>
<%
	end if ' blank company name
	connComp.MoveNext
Loop 

connComp.Close
Set connComp = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
				</table></td></tr>
		<tr><td></td><td><br><input type="submit" value="Add User"></td></tr>
	</form>
</table>
<%	End If ' userExists %>
</body>
</html>
