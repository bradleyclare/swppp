<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginAdmin.asp")
End If

inspecID = Request("inspecID")
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : Site Map Upload</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>SWPPP Inspections : Upload Site Map File</h1></td>
	</tr>
</table>
<% If Request.ServerVariables("request_method") <> "POST" Then ' Uppercase "POST" (Important) %>
<table width="100%" border="0">
<form action="<% = Request.ServerVariables("script_name") %>" enctype="multipart/form-data" method="post" onSubmit="return isReady(this)";>
		<input type="hidden" name="inspecID" value="<% = inspecID %>">
		<tr><td colspan="2" align="center">Upon completion, this page will redirect 
				to the report administration<br>
				for file association and immediate editing.</td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<tr><td width="25%" align="right" nowrap>Local File Name:</td>
			<td width="75%"><input name="localFile" type="file" size="50"></td></tr>
		<tr><td height="44" align="left">&nbsp;</td>
			<td height="44" valign="bottom"> <input type="submit" value="Start Upload"></td></tr>
	</form>
</table>
<% Else 
		Server.ScriptTimeout=4500 'default is 90 seconds %>
<table width="100%" border="0">
<%	baseDir = "d:\vol\swpppinspections.com\www\htdocs\images\"
	Set upLoad = Server.CreateObject("SoftArtisans.FileUp")
	upLoad.Path = baseDir & "sitemap\"
	' Added for duplicate file names
	upLoad.CreateNewFile = True	
	If upLoad.IsEmpty Then %>
	<tr><td>The file that you tried to upload was empty. Most likely, you did 
			not specify a valid filename to your browser or you left the filename 
			field blank. Please <a href="<% = Request.ServerVariables("script_name") %>">try 
			again to process</a> another image.</td></tr>
<%	ElseIf upLoad.ContentDisposition <> "form-data" Then 	%>
	<tr><td>Your upload did not succeed, most likely because your browser does 
			not support document upload via this mechanism. Please consult your 
			systems director and/or continue <a href="../">other administrative 
			functions</a> with this user session.</td></tr>
<%	Else
		On Error Resume Next
		' upLoad.Save		
		If Err <> 0 Then %>
	<tr><td><font color="red">An error occurred when saving the file on the server. 
			Possible causes include:</font>
			<ul><li>An incorrect filename was specified when trying to upload.</li>
				<li>File permissions do not allow writing to the specified area.</li></ul>
			Please contact <a href="http://www.swppp.com/">SWPPP</a> 
			for more troubleshooting information, or send e-mail to <a href="mailto:dwims@swppp.com">susan@plonka.com</A> 
			for additional help on this matter.</td>
</table>
<%		Else
			upLoad.Save			
			inspecID = upLoad.Form(1)			
			Set upLoad = Nothing			
			Response.Redirect("editReport.asp?inspecID=" & inspecID)			
		End If ' On Error		
	End If ' Is Empty	
	Set upLoad = Nothing	
End If ' Form Request %>
</body>
</html>