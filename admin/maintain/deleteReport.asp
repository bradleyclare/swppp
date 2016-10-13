<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") &_
		"?" & Request("query_string")
	Response.Redirect("loginUser.asp")
End If

%> <!-- #include file="../connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
'--	Delete all coordinates -------------------------------------------------------
	SQLDELETE = "DELETE FROM Coordinates WHERE inspecID=" & Request("inspecID")
'--	Delete all Images ------------------------------------------------------------	
	SQLDELETE = SQLDELETE &" DELETE FROM Images WHERE inspecID=" & Request("inspecID")
'--	Delete all Optional Image Relationships --------------------------------------
	SQLDELETE = SQLDELETE & " DELETE FROM OptionalImages WHERE inspecID=" &Request("inspecID")
'-- Delete the Inspection Report -------------------------------------------------
	SQLDELETE = SQLDELETE & " DELETE FROM Inspections WHERE inspecID=" & Request("inspecID")
	connSWPPP.Execute(SQLDELETE)	
	connSWPPP.Close
	Set connSWPPP = Nothing
	Response.Redirect("viewReports.asp")
end If
%>
<html>
<head>
	<title>SWPPP INSPECTIONS : Delete Report</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validCoordinates.js"></script>
	<script language="JavaScript" src="../js/validCoordinates1.2.js"></script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Delete Report</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
<form action="<%= Request.ServerVariables("script_name") %>" method="post">
	<input type="hidden" name="inspecID" value="<%= Request("inspecID") %>">
	<tr><td align="center"><font color="#FF0000">
		You are about to permanently remove this report.<br>
		This will also delete any related image and coordinate records.<br>
		You can return to <a href="../default.asp">Admin Home</a> to abort this operation.<br>
		Do you want to continue this delete process?</font>
		
		<br><br><input type="submit" value="Yes, Delete">
	</form>
</td></tr></table>
</body>
</html>
