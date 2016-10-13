<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

inspecID = Session("inspecID")
If Request.Form.Count > 0 Then
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function

	If Request("smallImage")="" OR Request("largeImage")="" then
		If Request("smallImage")="" then smallImageBlank=true end if
		If Request("largeImage")="" then largeImageBlank=true end if
	else
		SQLINSERT = "INSERT INTO Images (largeImage, smallImage, description, inspecID" & _
			") VALUES (" &_
			"'" & Request("largeImage") & "'" & _
			", '" & Request("smallImage") & "'" & _
			", '" & strQuoteReplace(Request("description")) & "'" & _
			", " & inspecID &_
			")"
'-- Response.Write(SQLINSERT & "<br>")
		connSWPPP.Execute(SQLINSERT)
	
		connSWPPP.Close
		Set connSWPPP = Nothing	
	end if
end if

%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
SQLSELECT = "SELECT projectName" &_
	" FROM Inspections" &_
	" WHERE inspecID=" & inspecID
Set connReport = connSWPPP.execute(SQLSELECT) %>
<html>
<head><title>SWPPP INSPECTIONS : Add Image Association</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validAddImage.js"></script></head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Add Image Association</h1><% 
If smallImageBlank then %>
	<font color="#FF0000">Small image is blank. Please select a small image.</font><br><br>
<% end if
If largeImageBlank then %>
	<font color="#FF0000">Large image is blank. Please select a large image.</font><br><br>
<% end if  %>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
<form method="post" action="<% = Request.ServerVariables("script_name") %>">
<input type="hidden" name="inspecID" value="<% = inspecID %>">
<tr><td align="right">Inspection Report:&nbsp;&nbsp;</td>
	<td><%= Trim(connReport("projectName")) %></td></tr>	
<tr><td align="right">Description:&nbsp;&nbsp;</td>
	<td><input type="text" name="description" size="40" maxlength="40"></td></tr>
<!--- ------------------------------- Small Image ---------------------------- -->	
<tr><td align="right">Small Image:&nbsp;&nbsp;</td>
	<td><SELECT name="smallImage"><%
' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") &"images\sm\")
Set gifDirectory = objFolder.Files

for each gifFile in gifDirectory
	slashLoc = InStrRev(gifFile,"\")
	pathLen = Len(gifFile)
	shortenedName = Right(gifFile,pathLen-slashLoc)
	SQL0="SELECT smallImage FROM Images" &_
		" WHERE smallImage='"& shortenedName &"'"
'Response.Write(SQL0&"<br>")
	SET RS0= connSWPPP.execute(SQL0)
	IF (RS0.BOF AND RS0.EOF) THEN %><option value="<%= shortenedName %>"><%= shortenedName %></option><% END IF	'-- Image is not attached --
next %></SELECT>
<% ' set return to for after upload
Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string") %>
		<input type="button" value="Upload Small Image" onClick="location='upSmImgRprt.asp?inspecID=<% = inspecID %>'; return false";></td></tr>
<!--- ----------------------------- Large Image --------------------------------- --->	
<tr><td align="right"><br>Large Image:&nbsp;&nbsp;</td>
	<td><br><select name="largeImage"><%
' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") &"images\lg\")
Set gifDirectory = objFolder.Files
'Response.Write(gifDirectory&"<br>")
for each gifFile in gifDirectory
'Response.Write(gifFile&"<br>")
	slashLoc = InStrRev(gifFile,"\")
	pathLen = Len(gifFile)
	shortenedName = Right(gifFile,pathLen-slashLoc) 
'Response.Write(shortenedName&"<br>")
	SQL1="SELECT largeImage FROM Images" &_
		" WHERE largeImage='"& shortenedName &"'"
	SET RS1= connSWPPP.execute(SQL1)
	IF (RS1.BOF AND RS1.EOF) THEN	'---- image is not in table ----------------------
%>	<option value="<%= shortenedName %>"><%= shortenedName %></OPTION>
<% 	END IF	'---- Image is not attached -----------------------------------------
 	next %>	
	</select>&nbsp;<input type="button" value="Upload Large Image" 
		onClick="location='upLgImgRpt.asp'; return false";></td></tr>	
	<tr><td colspan="2" align="center"><br><input type="submit" value="Add Association">
		<INPUT type="button" value="Return to Report View" onClick="window.navigate('editReport.asp?inspecID=<%= inspecID %>');"> 
	</td></tr>
</table><br><br>
</body>
</html><%
connSWPPP.Close 
Set connSWPPP = Nothing %>