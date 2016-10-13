<% 
' Get text from a file and inserts into page

If not Session("validAdmin") then
	Session("adminReturnTo") = Request.ServerVariables("PATH_INFO")
 	Response.Redirect("../maintain/loginAdmin.asp")
end if

base_path = server.mappath("../..")

' iomode settings
ForReading = 1
ForWriting = 2
ForAppending = 8

'format settings
TristateUseDefault = -2
TristateTrue = -1
TristateFalse = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Request.Form.Count > 0 Then
	'Response.write(base_path & "/static/contact.txt")
	objFSO.CreateTextFile base_path & "/static/contact.txt"
	Set objF = objFSO.GetFile(base_path & "/static/contact.txt")
	Set objFile = objF.OpenAsTextStream(ForWriting, TristateUseDefault)
	objFile.Write(Request.Form("content"))
	Response.Redirect("../../static/contact.asp")
end if

'response.write(base_path & "/static/contact.txt")
Set objFile = objFSO.OpenTextFile(base_path & "/static/contact.txt")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP INSPECTIONS: Admin : Edit Contact</title>
	<LINK REL=stylesheet HREF="../../global.css" TYPE="text/css">
</head>

<!-- #include file="../adminHeader2.inc" -->
<h1>Contact Us</h1>

<form action="<% = Request.ServerVariables("script_name") %>" method="POST">
	<textarea cols="70" rows="20" name="content"><%= objFile.ReadAll %></textarea><br><br>
	<input type="Submit" value="Publish">&nbsp;<input type="Reset">
</form></body>
</html>
<% objFile.Close %>