<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginAdmin.asp")
End If %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : Maintain : Archive Project Review</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>SWPPP Inspections : Maintain : Archive Project Review</h1></td>
	</tr>
</table>
This page will do a side by side comparison of the archive files to help the user verify that all files were created correctly.
</body>
</html>