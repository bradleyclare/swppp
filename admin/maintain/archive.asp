<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../maintain/loginUser.asp")
End If 

base_path = server.mappath(".")

'-- delete all subfolders and files from the temp_archives directory
DIM fso, f, fl, fc
SET fso = CreateObject("Scripting.FileSystemObject")
localDest = base_path & "\temporary_archives\"
Set f = fso.GetFolder(localDest)
Set fc = f.SubFolders
For Each fl in fc
	fl.Delete
Next

SQL0="sp_GetAllProjects"
%><!-- #include virtual="admin/connSWPPP.asp" --><%
SET RS0=connSWPPP.Execute(SQL0) %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SWPPP INSPECTIONS : Maintain : Archive</title>
    <link rel="stylesheet" href="../../global.css" type="text/css">
    <script language="JavaScript" src="../js/validUpload.js"></script>
    <script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>SWPPP Inspections : Maintain : Archive</h1></td>
	</tr>
</table>
This page will be used to select the Project to Archive.<br>
<!--<%'-- IF Request("err")=001 THEN %><FONT color="red">You Must enter a complete valid directory for your local machine ending in '\'.</FONT><%'-- END IF %>-->
<% IF Request("err")=002 THEN %><FONT color="red">You Must select a project for Archival</FONT><% END IF %>
<form action="archivePreview.asp" method="post">
<table>
	<tr><td><select name="projectID">
<% 	DO WHILE NOT RS0.EOF %><option value="<%=RS0(0)%>"><%=RS0(1)%>&nbsp;<%=RS0(2)%>
<%		RS0.MoveNext
	LOOP %>
			</select></td>
		<td valign="middle" align="center"><input type="submit" value="Submit">
	</tr>
</table>
</form>
</body>
</html>
<% SET RS0=nothing
connSWPPP.close() %>