<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") &_
		"?" & Request("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID=Session("inspecID")

%> <!-- #include file="../connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	SQLDELETE = "DELETE FROM Coordinates WHERE coID=" & Request("coID")
	' Response.Write(coordSQLINSERT & "<br><br>")
	connSWPPP.Execute(SQLDELETE)
	
	connSWPPP.Close
	Set connSWPPP = Nothing
	Response.Redirect("editReport.asp")
End If
%>
<html>
<head>
	<title>SWPPP INSPECTIONS : Delete Coordinate</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validCoordinates.js"></script>
	<script language="JavaScript" src="../js/validCoordinates1.2.js"></script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Delete Coordinate</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
<form action="<%= Request.ServerVariables("script_name") %>" method="post">
	<input type="hidden" name="coID" value="<%= Request("coID") %>">
	<tr><td align="center"><font color="#FF0000">You are about to permanently remove this record.<br>
	You can return to <a href="../default.asp">Admin Home</a> to abort this operation.<br>
	Do you want to continue this delete process?</font>
	
	<br><br><input type="submit" value="Yes, Delete"></td></tr>
	</form>
</table>
</body>
</html>
