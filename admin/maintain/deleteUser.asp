<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") &_
		"?" & Request("query_string")
	Response.Redirect("loginUser.asp")
End If

%> <!-- #include file="../connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	SQLSELECT = "SELECT inspecID, projectName" &_
		" FROM Inspections" &_
		" WHERE userID=" & Request("IDuser")
	Set connUser = connSWPPP.execute(SQLSELECT)
	
	If not connUser.eof then
		inspectionFound=true
	else
		SQLDELETE = "DELETE FROM CompanyUsers WHERE userID=" & Request("IDuser")
		' Response.Write(coordSQLINSERT & "<br><br>")
		connSWPPP.Execute(SQLDELETE)

		SQLDELETE = "DELETE FROM ProjectsUsers WHERE userID=" & Request("IDuser")
		' Response.Write(coordSQLINSERT & "<br><br>")
		connSWPPP.Execute(SQLDELETE)
		
		SQLDELETE = "DELETE FROM Users WHERE userID=" & Request("IDuser")
		' Response.Write(coordSQLINSERT & "<br><br>")
		connSWPPP.Execute(SQLDELETE)
		
		connSWPPP.Close
		Set connSWPPP = Nothing
		Response.Redirect("viewUsersAdmin.asp")
	end if
end If
%>
<html>
<head>
	<title>SWPPP Inspections : Delete User</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validCoordinates.js"></script>
	<script language="JavaScript" src="../js/validCoordinates1.2.js"></script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Delete User</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
<form action="<%= Request.ServerVariables("script_name") %>" method="post">
	<input type="hidden" name="IDuser" value="<%= Request("IDuser") %>">
	<tr><td align="center"><font color="#FF0000">
	<% If inspectionFound then %>
		A report exists with this user as the inspector of record.<br>
		The user must not be associates with an existing report to be deleted.<br>
		<a href="viewReports.asp">View Reports</a>
	<% else %>
		You are about to permanently remove this record.<br>
		You can return to <a href="../default.asp">Admin Home</a> to abort this operation.<br>
		Do you want to continue this delete process?</font>
		
		<br><br><input type="submit" value="Yes, Delete">
	<% end if %>
	</form>
</td></tr></table>
</body>
</html>
