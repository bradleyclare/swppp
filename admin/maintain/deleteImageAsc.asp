<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID = Session("inspecID")
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	SQLDELETE = "DELETE FROM Images WHERE imageID=" & Request("imageID")
	Response.Write(SQLDELETE & "<br>")
	connSWPPP.Execute(SQLDELETE)
	
	connSWPPP.Close
	Set connSWPPP = Nothing
	
	Response.Redirect("editReport.asp")
end if
%>

<html>
<head>
	<title>SWPPP INSPECTIONS : Admin: Delete Image Association</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validAddImage.js"></script>
</head>

<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Delete Image Association</h1>
<table><tr><td align="center">
	<font color="red">
	You are about to delete the association for this image.<br>
	This will not delete the image only the association of the image with the report.<br>
	To abort this operation you can return to <a href="viewReports.asp">View Reports</a>. 
	</font>
	<form method="post" action="<%= Request.ServerVariables("script_name") %>">
	<input type="hidden" name="imageID" value="<%= Request("imageID") %>">
	<input type="submit" value="Delete Association">
	</form>
</td></tr></table><br><br>
</body>
</html>
