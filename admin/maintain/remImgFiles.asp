<%@ Language="VBScript" %>
<%
If not Session("validAdmin") then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
 	Response.Redirect("loginUser.asp")
end if %>
<!-- #include file="../connSWPPP.asp" --><%
IF TRIM(Request("oitID"))<>"" THEN oitID=Request("oitID") ELSE oitID=1 END IF
IF NOT(IsNumeric(oitID)) THEN oitID=1 END IF 
oitID=CINT(oitID)
SQL0="SELECT MIN(oitID) as minVal, MAX(oitID) as maxVal FROM OptionalImagesTypes"
SET RS0=connSWPPP.execute(SQL0)
IF (oitID<RS0("minVal") OR oitID>RS0("maxVal")) THEN
	optImage=False
	SELECT CASE oitID
		CASE -1
			imgDir="lg"
			folderName=imgDir
			columnName="largeImage"
			imgType="Large Image"
	        baseDir = Server.Mappath("\images\" & folderName )
		CASE else
			oitID=-2
			imgDir="sm"
			folderName=imgDir
			columnName="smallImage"
			imgType="Small Image"
	        baseDir = Server.Mappath("\images\" & folderName )
	END SELECT
ELSE
	SQL1="SELECT * FROM OptionalImagesTypes WHERE oitID="& oitID
	SET RS1=connSWPPP.execute(SQL1)
	optImage=True
	imgDir=Trim(RS1("oitName"))
	folderName=imgDir
	imgType=Trim(RS1("oitDesc"))
	baseDir = Server.Mappath("\images\" & folderName )
END IF
If Request.Form.Count > 0 Then
	iRemCount = Request("iRemCount")
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	For i = 1 To iRemCount
		remImgFile = Request("remImgFile" & i)	
		If remImgFile <> "" Then folderSvrObj.DeleteFile(baseDir &"\"& remImgFile) End If
	Next		
	Set folderSvrObj = Nothing		
End If %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS <%= imgType %> : Image : Removal</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>

<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><h1>Remove Images <%= imgType %></h1></td></tr></table>
<% If Request("list") Then
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")	
	Set objImageDir = folderSvrObj.GetFolder(baseDir &"\")
	Set remImgFiles = objImageDir.Files %>
<table width="100%" border="0">
	<tr><td align="center">Please select image files not associated in database for removal.</td></tr>
	<tr><td>&nbsp;</td></tr></table>	
<table width="100%" border="0">
<form method="post" action="<% = Request.ServerVariables("script_name") %>">
	<input type="hidden" name="oitID" value="<% = oitID %>">
	<tr><th width="40%"><b>File Name</b></th>
	    <th width="45%"><b>File Date</b></th>
	    <th width="5%"><b>Select</b></th></tr>
	<tr><td colspan="4"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
<!-- #include file="../connSWPPP.asp" -->
<%	iRemCount = 0
	altColors = "#e5e6e8"
	IF remImgFiles.count>0 THEN
		For Each Item In remImgFiles
			iFileName = Item.Name
			iFileDate = Item.DateCreated
			IF optImage THEN 
				imgSQLSELECT = "SELECT * FROM OptionalImages WHERE oImageFileName='"& iFileName &"' AND oitID='"& oitID &"'"
			ELSE 
				imgSQLSELECT = "SELECT imageID FROM Images WHERE " & columnName & " = " & "'" & iFileName & "'" 
			END IF
			Set rsImage = connSWPPP.Execute(imgSQLSELECT)
			IF rsImage.EOF THEN
				iRemCount = iRemCount + 1 %>
	<tr bgcolor="<% = altColors %>" align="center" nowrap> 
		<td><% = iFileName %></td>
		<td><% = iFileDate %></td>
	    <td><input type="checkbox" name="remImgFile<% = iRemCount %>" value="<% = iFileName %>"></td>
	</tr>
<%				If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			End If
		Next
		rsImage.Close
		Set rsImage = Nothing
	ELSE %>
	<tr bgcolor="<% = altColors %>" align="center" nowrap>There are no Images in the Directory that you are searching</td></tr>
<%	END IF
	Set folderSvrObj = Nothing
	Set objImageDir = Nothing
	Set remImgFiles = Nothing	
	If iRemCount <> 0 Then %>
	<tr align="center"> 
		<td colspan="4">&nbsp;<br>
			<input type="submit" value="Remove"><br>
		&nbsp;</td>
	</tr>
<%	Else
		Response.Write("<tr><td colspan='4' align='center'>&nbsp;<br><b><i>Sorry " & _
			"no current unassociated images at this time.</i></b><br>&nbsp;</td></tr>")
	End If %>
	<input type="hidden" name="iRemCount" value="<% = iRemCount %>">
</form>
</table>
<% Else 
	SQL0="SELECT * FROM OptionalImagesTypes ORDER BY oitSortByVal asc"
	SET RS0=connSWPPP.execute(SQL0)	%>
<table width="100%" border="0">
	<tr> 
		<td align="center">Please select an image size needed to be listed for removal.</td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
	</tr>
	<tr><td><ul>
				<li><a href="remImgFiles.asp?oitID=-2&list=True">Small</a><br>&nbsp;</li>
				<li><a href="remImgFiles.asp?oitID=-1&list=True">Large</a><br>&nbsp;</li><%
DO WHILE NOT RS0.EOF %>
				<li><a href="remImgFiles.asp?oitID=<%= RS0("oitID")%>&list=True"><%= TRIM(RS0("oitDesc"))%></a><br>&nbsp;</li><%
	RS0.MoveNext
LOOP %></ul></td>
	</tr>
</table>
<% End If 
	connSWPPP.Close
	Set connSWPPP = Nothing %>
</body>
</html>