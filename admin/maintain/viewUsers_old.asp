<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

recordOrd = Request("orderBy")
If recordOrd = "" Then
	recordOrd = "lastName"
End If

%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
If Session("validDirector") then
	'select the companies for which this user is a valid Director
	SQLSELECT = "SELECT companyID" & _
			" FROM CompanyUsers" &_
			" WHERE userID=" & Session("userID") &_
			" AND rights='director'"
	Set connComp = connSWPPP.Execute(SQLSELECT)
	
	SQLSELECT = "SELECT DISTINCT Users.userID, firstName, lastName" & _
		" FROM Users, CompanyUsers, Companies" & _
		" WHERE Users.userID = CompanyUsers.userID" &_
		" AND (rights='director' OR rights='user')" &_
		" AND CompanyUsers.companyID=Companies.companyID"
	Do while not connComp.eof
		if not subsequent then 'first time
			SQLSELECT = SQLSELECT & " AND (CompanyUsers.companyID=" & connComp("companyID")
			subsequent=true
		else
			SQLSELECT = SQLSELECT & " OR CompanyUsers.companyID=" & connComp("companyID")
		end if
		connComp.movenext
	Loop

	SQLSELECT = SQLSELECT & ") ORDER BY " & recordOrd
	Response.Write(SQLSELECT & "<br>")
	connComp.close
else 'validAdmin
	SQLSELECT = "SELECT userID, firstName, lastName" &_
		" FROM Users" &_
		" ORDER BY " & recordOrd
	'Response.Write(SQLSELECT & "<br>")
end if

Set connUsers = connSWPPP.Execute(SQLSELECT)

recCount = 0
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : View Users</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><br><h1>View Users</h1></td></tr></table>
<table width="100%" border="0">
	<tr><th><b>Count</b></th>
		<th><a href="<% = Request.ServerVariables("script_name") %>?orderBy=firstName"> 
			<b>First Name</b></a></th>
		<th><a href="<% = Request.ServerVariables("script_name") %>?orderBy=lastName"> 
			<b>Last Name</b></a></th>
		<th><b>Company</b></th>
		<th><b>Rights</b></th></tr>
<%
	If connUsers.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no users.</i></b></td></tr>")
	Else
		altColors="#ffffff"
		
		Do While Not connUsers.EOF
			recCount = recCount + 1
			SQLSELECT =	"SELECT CompanyUsers.companyID, rights, companyName" &_
				" FROM CompanyUsers, Companies" &_
				" WHERE CompanyUsers.userID=" & connUsers("userID") &_
				" AND CompanyUsers.companyID=Companies.companyID" &_
				" ORDER BY rights"
			Set connRights=connSWPPP.execute(SQLSELECT)
			If connRights.eof then %>
	<tr align="center" bgcolor="<%= altColors %>"> 
		<td><%= recCount %></td>
		<td><%= Trim(connUsers("firstName")) %></td>
		<td><a href="editUser.asp?IDuser=<%= connUsers("userID") %>">
			<%= Trim(connUsers("lastName")) %></a></td></tr>
<%
			else
				Do While not connRights.eof
%>
	<tr align="center" bgcolor="<%= altColors %>"> 
		<td><%= recCount %></td>
		<td><%= Trim(connUsers("firstName")) %></td>
		<td><a href="editUser.asp?IDuser=<%= connUsers("userID") %>">
			<%= Trim(connUsers("lastName")) %></a></td>
		<td nowrap><%= connRights("companyName") %></td>
		<td><%= connRights("rights") %></td></tr>
<%
					connRights.movenext
				Loop
			end if 'any rights for this user?
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
