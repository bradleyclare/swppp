<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : Maintain : Archive Completed</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>SWPPP Inspections : Maintain : Archive Completed</h1></td>
	</tr>
</table>
This page will show the final file structure created for review by the User and delete the Temporary records.
</body>
</html>