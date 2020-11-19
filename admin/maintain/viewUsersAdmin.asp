<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

If Session("validDirector") then
	Response.Redirect("viewUsersDir.asp")
end if

recordOrd = Request("orderBy")
If recordOrd = "" Then
	recordOrd = "lastName"
End If

%> <!-- #include file="../connSWPPP.asp" --> <%
SQLSELECT = "SELECT userID, firstName, lastName, rights" &_
	" FROM Users WHERE active=1" &_
	" ORDER BY " & recordOrd
'Response.Write(SQLSELECT & "<br>")
Set connUsers = connSWPPP.Execute(SQLSELECT)

recCount = 0
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : View Users for Admins</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><br><h1>View Users</h1></td></tr></table>
<table width="100%" border="0">
	<tr><th><b>Count</b></th>
		<th><a href="<% = Request.ServerVariables("script_name") %>?orderBy=firstName"> 
			<b>First Name</b></a></th>
		<th><a href="<% = Request.ServerVariables("script_name") %>?orderBy=lastName"> 
			<b>Last Name</b></a></th>
		<th><a href="<% = Request.ServerVariables("script_name") %>?orderBy=rights">
		    <b>Rights</b></a></th></tr>
<%
	If connUsers.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no users.</i></b></td></tr>")
	Else
		altColors="#ffffff"
		
		Do While Not connUsers.EOF
			recCount = recCount + 1 %>
	<tr align="center" bgcolor="<%= altColors %>"> 
		<td><%= recCount %></td>
		<td><%= Trim(connUsers("firstName")) %></td>
		<td><a href="editUser.asp?IDuser=<%= connUsers("userID") %>">
			<%= Trim(connUsers("lastName")) %></a></td>
		<td><%= connUsers("rights") %></td></tr>
<%			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
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
