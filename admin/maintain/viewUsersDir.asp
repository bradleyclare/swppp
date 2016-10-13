<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") AND not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

If Session("validAdmin") then
	Response.Redirect("viewUsersAdmin.asp")
end if

recordOrd = Request("orderBy")
If recordOrd = "" Then
	recordOrd = "lastName"
End If

%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
'select the companies for which this user is a valid Director
SQLSELECT = "SELECT projectID" & _
		" FROM ProjectsUsers" &_
		" WHERE userID=" & Session("userID") &_
		" AND rights='director'"
Set connComp = connSWPPP.Execute(SQLSELECT)

' select users who have rights to those companies
SQLSELECT = "SELECT DISTINCT u.userID, firstName, lastName, pu.rights" &_
	" FROM Users as u JOIN ProjectsUsers as pu" &_
	" ON u.userID=pu.userID JOIN Projects as p" &_
	" ON pu.projectId=p.projectID" &_
	" WHERE u.userID = pu.userID AND pu.rights IN ('director', 'user')"  &_
	" AND u.rights!='admin' AND p.projectID IN (" 
Do while not connComp.eof
	if not subsequent then 'first time
		SQLSELECT = SQLSELECT & connComp("projectID")
		subsequent=true
	else
		SQLSELECT = SQLSELECT & ", " & connComp("projectID")
	end if
	connComp.movenext
Loop

SQLSELECT = SQLSELECT & ") ORDER BY " & recordOrd
'Response.Write(SQLSELECT & "<br>")
connComp.movefirst
Set connUsers = connSWPPP.Execute(SQLSELECT)

connComp.Close
Set connComp = Nothing
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : View Users for Directors</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><br><h1>View Users</h1></td></tr></table>
<table width="100%" border="0">
	<tr><th><b>Count</b></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=firstName"> 
			<b>First Name</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=lastName"> 
			<b>Last Name</b></a></th>
		<th><b>Rights</b></th></tr>
<%
	If connUsers.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no users.</i></b></td></tr>")
	Else
		altColors="#ffffff"
		recCount = 0
		Do While Not connUsers.EOF
			recCount=recCount+1
%>
	<tr align="center" bgcolor="<%= altColors %>"> 
		<td><%= recCount %></td>
		<td><%= Trim(connUsers("firstName")) %></td>
		<td><a href="editUser.asp?IDuser=<%= connUsers("userID") %>">
			<%= Trim(connUsers("lastName")) %></a></td>
		<td><%= connUsers("rights") %></td></tr>
<%
			' Alternate Row Colors
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			connUsers.MoveNext
		Loop
	End If ' END No Results Found

connUsers.Close
Set connUsers = Nothing
connSWPPP.Close
Set connSWPPP = Nothing
%>
</table>
</body>
</html>
