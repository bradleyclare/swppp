<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If 

'-- delete all subfolders and files from the temp_archives directory
DIM fso, f, fl, fc
SET fso = CreateObject("Scripting.FileSystemObject")
localDest = Request.ServerVariables("APPL_PHYSICAL_PATH") & "admin\maintain\temporary_archives\"
Set f = fso.GetFolder(localDest)
Set fc = f.SubFolders
For Each fl in fc
    Response.Write("Attemping to Delete: " & fl)
	fl.Delete
Next

SQL0="sp_GetAllProjects"
%><!-- #include file="../connSWPPP.asp" --><%
SET RS0=connSWPPP.Execute(SQL0) %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : Maintain : Archive</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>SWPPP Inspections : Maintain : Archive</h1></td>
	</tr>
</table>
This page will be used to select the Project to Archive.<br>
<!--<%'-- IF Request("err")=001 THEN %><FONT color="red">You Must enter a complete valid directory for your local machine ending in '\'.</FONT><%'-- END IF %>-->
<% IF Request("err")=002 THEN %><FONT color="red">You Must select a project for Archival</FONT><% END IF %>
<FORM action="archivePreview.asp" method="post">
<TABLE>
	<TR><TD><SELECT name="projectID">
<% 	DO WHILE NOT RS0.EOF %><OPTION value="<%=RS0(0)%>"><%=RS0(1)%>&nbsp;<%=RS0(2)%>
<%		RS0.MoveNext
	LOOP %>
			</SELECT></TD>
		<TD valign="middle" align="center"><input type="submit" value="Submit">
	</TR>
</TABLE>
</body>
</html>
<% SET RS0=nothing
connSWPPP.close() %>