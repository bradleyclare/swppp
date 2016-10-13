<%@ Language="VBScript" %>
<%
If not Session("validAdmin") then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
 	Response.Redirect("loginUser.asp")
end if

folderName = "sitemap"
displaySize = " : Site Maps"

If Request.Form.Count > 0 Then
	mRemCount = Request("mRemCount")
	
	baseDir = "d:\vol\swpppinspections.com\www\htdocs\images\" & folderName & "\"
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	
	For i = 1 To mRemCount
		remMapFile = Request("remMapFile" & i)
		
		' Response.Write(remMapFile & "<br>")
		If remMapFile <> "" Then folderSvrObj.DeleteFile(baseDir & remMapFile) End If
		
	Next
	
	Set folderSvrObj = Nothing
	
	Response.Redirect(Request.ServerVariables("path_info"))
	
End If
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS<% = displaySize %> : Removal</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>Remove<% = displaySize %></h1></td>
	</tr>
</table>
<%
baseDir = "d:\vol\swpppinspections.com\www\htdocs\images\" & folderName
Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")

Set objImageDir = folderSvrObj.GetFolder(baseDir)
Set remMapFiles = objImageDir.Files
%>
<table width="100%" border="0">
	<tr>
		<td align="center">Please select site map files not associated in database 
			for removal.</td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
	</tr>
</table>
<table width="100%" border="0">
<form method="post" action="<% = Request.ServerVariables("script_name") %>">
	<tr> 
		    <th width="20%"><b>Site Map</b></th>
		    <th width="35%"><b>File Name</b></th>
		    <th width="35%"><b>File Date</b></th>
		    <th width="10%"><b>Select</b></th>
	</tr>
	<tr> 
		<td colspan="4"><img src="../../images/dot.gif" width="5" height="5"></td>
	</tr>
<!-- #include virtual="admin/connSWPPP.asp" -->
<%
mRemCount = 0

altColors = "#e5e6e8"

For Each Item In remMapFiles
	mFileName = Item.Name
	mFileDate = Item.DateCreated
	
	mapSQLSELECT = "SELECT inspecID FROM OptionalImages " & _
		" WHERE oImageType='sitemap' AND oImageFileName='"& mFileName &"'"
		
	' Response.Write("mapSQLSELECT")
	Set rsMapFile = connSWPPP.Execute(mapSQLSELECT)
	
	If rsMapFile.EOF Then
		
		mRemCount = mRemCount + 1
%>
	<tr bgcolor="<% = altColors %>" align="center" nowrap> 
		    <td><a href="../../images/<% = folderName & "/" & mFileName %>" target="_blank"><% = mFileName %></a></td>
		<td> 
			<% = mFileName %>
		</td>
		<td> 
			<% = mFileDate %>
		</td>
	    <td><input type="checkbox" name="remMapFile<% = mRemCount %>" value="<% = mFileName %>"></td>
	</tr>
<%
		If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
		
		rsMapFile.Close
		Set rsMapFile = Nothing
		
	End If
	
Next

connSWPPP.Close
Set connSWPPP = Nothing

Set folderSvrObj = Nothing
Set objImageDir = Nothing
Set remMapFiles = Nothing

If mRemCount <> 0 Then
%>
	<tr align="center"> 
		<td colspan="4">&nbsp;<br>
			<input type="submit" value="Remove"><br>
		&nbsp;</td>
	</tr>
<%
Else
	Response.Write("<tr><td colspan='4' align='center'>&nbsp;<br><b><i>Sorry no " & _
		"current unassociated site maps at this time.</i></b><br>&nbsp;</td></tr>")
End If
%>
	<input type="hidden" name="mRemCount" value="<% = mRemCount %>">
</form>
</table>
</body>
</html>
