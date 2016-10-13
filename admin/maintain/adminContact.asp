<% 
' Get text from a file and inserts into page

If not Session("validAdmin") then
	Session("adminReturnTo") = Request.ServerVariables("PATH_INFO")
 	Response.Redirect("../maintain/loginUser.asp")
end if

base_path = server.mappath("/")

' iomode settings
ForReading = 1
ForWriting = 2
ForAppending = 8

'format settings
TristateUseDefault = -2
TristateTrue = -1
TristateFalse = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")

updateFlag = False
If Request.Form.Count > 0 Then
	'response.Write(base_path & "/contact.txt")
	Set objF = objFSO.CreateTextFile(base_path & "/contact.txt",True)
	'Set objF = objFSO.GetFile(base_path & "/contact.txt")
	'Set objFile = objF.OpenAsTextStream(ForWriting, TristateUseDefault)
	objF.Write(Request.Form("content"))
	'Response.Redirect(base_path & "/contact.txt")
	objF.Close
	updateFlag = True
end if

'response.write(base_path & "/contact.txt")
Set objFile = objFSO.OpenTextFile(base_path & "/contact.txt")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP INSPECTIONS: Admin : Edit Contact</title>
	<LINK REL=stylesheet HREF="../../global.css" TYPE="text/css">
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Contact Us</h1>

    <form action="<% = Request.ServerVariables("script_name") %>" method="POST">
	    <textarea cols="70" rows="10" name="content"><%= objFile.ReadAll %></textarea><br><br>
	    <input type="Submit" value="Publish">&nbsp;<input type="Reset">
    </form>
</body>
<!-- TinyMCE --> 
<script type="text/javascript" src="../../js/tinymce/tinymce.min.js"></script>
<script type="text/javascript">
tinymce.init({
	selector: "textarea",
	plugins: [
		"advlist autolink lists link image charmap print preview anchor",
		"searchreplace visualblocks code fullscreen",
		"insertdatetime media table contextmenu paste"
	],
	toolbar: "insertfile undo redo | styleselect | bold italic | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | link image"
});
</script>
<!-- /TinyMCE -->
</html>
<% objFile.Close %>