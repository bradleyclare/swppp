<%@ Language="VBScript" %>
<%
'also for use by upload image
Session("adminReturnTo") = Request.ServerVariables("path_info") & "?imageID=" & Request("imageID")

If Not Session("validAdmin") And Not Session("validInspector") Then
	Response.Redirect("loginUser.asp")
End If
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	SQLUPDATE = "UPDATE Images SET" & _
		" largeImage='" & Request("largeImage") & "'" & _
		", smallImage='" & Request("smallImage") & "'" & _
		", description='" & strQuoteReplace(Request("description")) & "'" & _
		" WHERE imageID=" & Request("imageID")
	Response.Write(SQLUPDATE & "<br><br>")
	connSWPPP.Execute(SQLUPDATE)
	connSWPPP.Close
	Set connSWPPP = Nothing
	Response.Redirect("editReport.asp?inspecID="& Session("inspecID"))
else
	SQLSELECT = "SELECT largeImage, smallImage, description, projectName, Images.inspecID" &_
		" FROM Images, Inspections" &_
		" WHERE Images.imageID=" & Request("imageID") &_
		" AND Inspections.inspecID=Images.inspecID"
	'Response.Write(SQLSELECT & "<br>")
	Set connImage = connSWPPP.execute(SQLSELECT)
End If
inspecID=Session("inspecID") %>
<html>
<head>
	<title>SWPPP INSPECTIONS : Edit Image Association</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validReports.js"></script>
	<script language="JavaScript" src="../js/validReports1.2.js"></script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Edit Image Association</h1>	
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
<form method="post" action="<% = Request.ServerVariables("script_name") %>">
<input type="hidden" name="imageID" value="<%= Request("imageID") %>">
<input type="hidden" name="inspecID" value="<%= connImage("inspecID") %>">
<tr><td></td>
	<td><input type="button" value="Delete Image Association" 
	onClick="location='deleteImageAsc.asp?imageID=<%= Request("imageID") %>'; return false";><br><br></td></tr>
<tr><td align="right">Inspection Report:&nbsp;&nbsp;</td>
	<td><%= connImage("projectName") %></td></tr>	
<tr><td align="right">Description:&nbsp;&nbsp;</td>
	<td><input type="text" name="description" size="40" maxlength="40" 
		value="<%= Trim(connImage("description")) %>"></td></tr>
<tr><td align="right">Small Image:&nbsp;&nbsp;</td>
	<td><input type="button" value="Upload Small Image" 
		onClick="location='upSmImgRpt.asp'; return false";></td></tr>
<tr><td colspan=2 align="center">Current Image:<br><%= connImage("smallImage") %><br>
		<img src="<%= "../../images/sm/" & connImage("smallImage") %>" border="0" 
		alt="<%= connImage("smallImage") %>"><br>
		<input type="radio" name="smallImage" value="<%= connImage("smallImage") %>" checked></td></tr>
<tr><%
' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder("d:\vol\swpppinspections.com\www\htdocs\images\sm\")
Set gifDirectory = objFolder.Files

for each gifFile in gifDirectory
	slashLoc = InStrRev(gifFile,"\")
	pathLen = Len(gifFile)
	shortenedName = Right(gifFile,pathLen-slashLoc) 
	SQL0="SELECT smallImage FROM Images" &_
		" WHERE smallImage='"& shortenedName &"'"
'Response.Write(SQL0&"<br>")
	SET RS0= connSWPPP.execute(SQL0)
	IF (RS0.BOF AND RS0.EOF) THEN	'---- image is not in table ----------------------
	iDataRows = iDataRows + 1
	If iDataRows > 3 Then
		Response.Write("</tr>" & VBCrLf & "<tr>")
		iDataRows = 1
	End If 	%>
	<td align="center"><br><%= shortenedName %><br>
		<img src="<%= "../../images/sm/" & shortenedName %>" border="0" 
		alt="<%= gifFile %>"><br>
		<input type="radio" name="smallImage" value="<%= shortenedName %>"
		<% If Trim(connImage("smallImage"))=shortenedName then %>checked<% end if %>></td>
<% 	END IF	'---- Image is not attached -----------------------------------------
	next %>
	</tr>	
<tr><td align="right"><br>Large Image:&nbsp;&nbsp;</td>
	<td><br><select name="largeImage"><%
' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
'-Set objFolder = folderServerObj.GetFolder("d:\vol\swpppinspections.com\www\htdocs\images\lg\")		'-- live site
Set objFolder = folderServerObj.GetFolder("d:\vol\swpppinspections.com\www\htdocs\images\lg\")	'-- dev site
Set gifDirectory = objFolder.Files
for each gifFile in gifDirectory
	slashLoc = InStrRev(gifFile,"\")
	pathLen = Len(gifFile)
	shortenedName = Right(gifFile,pathLen-slashLoc) %>
	<option value="<%= shortenedName %>"<% If Trim(connImage("largeImage"))=shortenedName then %> selected<% end if %>><%= shortenedName %></OPTION><% 
next %>	
	</select>&nbsp;<input type="button" value="Upload Large Image" 
		onClick="location='upLgImgRpt.asp'; return false";></td></tr>
	
<tr><td colspan="2" align="center"><br><input type="submit" value="Modify Association"></td></tr>

</table><br><br>
</body>
</html><%
connSWPPP.Close 
Set connSWPPP = Nothing %>