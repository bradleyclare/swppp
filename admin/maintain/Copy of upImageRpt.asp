<%@ Language="VBScript" %>
<!-- #include file="../connSWPPP.asp" --> <%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginAdmin.asp")
End If 
inspecID = Request("inspecID") 
SQL0="SELECT MIN(oitID) as minVal, MAX(oitID) as maxVal FROM OptionalImagesTypes"
SET RS0=connSWPPP.execute(SQL0)
IF NOT(IsEmpty(Session("oitID"))) THEN oitID= Session("oitID") END IF
IF TRIM(Request("oitID"))<>"" THEN oitID=Request("oitID") END IF
IF NOT(IsNumeric(oitID)) THEN oitID=1 END IF 
oitID=CINT(oitID)
IF (oitID<RS0("minVal") OR oitID>RS0("maxVal")) THEN oitID=1 END IF
SQL1="SELECT * FROM OptionalImagesTypes WHERE oitID="& oitID
SET RS1=connSWPPP.execute(SQL1)
	Session("oitID")=oitID
	imgDir=Trim(RS1("oitName"))
	imgType=Trim(RS1("oitDesc"))
	baseDir = "d:\vol\swpppinspections.com\www\htdocs\images\" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : <%= imgType %> Upload</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><h1>SWPPP Inspections : Upload <%= imgType %> File</h1></td></tr>
</table>
<% If Request.ServerVariables("request_method") <> "POST" Then ' Uppercase "POST" (Important) %>
<table width="100%" border="0">
	<form action="<% = Request.ServerVariables("script_name") %>" enctype="multipart/form-data" method="post" onSubmit="return isReady(this)";>
		<input type="hidden" name="inspecID" value="<%= inspecID %>">
		<input type="hidden" name="oitID" value="<%= oitID %>">
		<tr><td colspan="2" align="center">Upon completion, this page will redirect 
				to the report administration<br>for file association and immediate editing.</td></tr>
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
<%	Set upLoad = Server.CreateObject("SoftArtisans.FileUp")
	upLoad.Path = baseDir & imgDir
	' Added for duplicate file names
	upLoad.CreateNewFile = True	
	If upLoad.IsEmpty Then %>
	<tr><td>The file that you tried to upload was empty. Most likely, you did 
			not specify a valid filename to your browser or you left the filename 
			field blank. Please <a href="<% = Request.ServerVariables("script_name") %>">try 
			again to process</a> another image.</td></tr>
<%	ElseIf upLoad.ContentDisposition <> "form-data" Then %>
	<tr><td>Your upload did not succeed, most likely because your browser does 
			not support document upload via this mechanism. Please consult your 
			systems director and/or continue <a href="../">other administrative 
			functions</a> with this user session.</td></tr>
<%	Else
		On Error Resume Next
		upLoad.Save
		If Err <> 0 Then %>
	<tr><td><font color="red">An error occurred when saving the file on the server. 
			Possible causes include:</font>
			<ul><li>An incorrect filename was specified when trying to upload.</li>
				<li>File permissions do not allow writing to the specified area.</li></ul>
			Please contact <a href="http://www.swppp.com/">SWPPP</a> 
			for more troubleshooting information, or send e-mail to <a href="mailto:dwims@swppp.com">susan@plonka.com</A> 
			for additional help on this matter.</td></tr>
</table>
<%		Else
			inspecID = upLoad.Form(1)			
			Set upLoad = Nothing
			If inspecID="" Then
				Response.Redirect("newReport.asp")
			Else		
				Response.Redirect("editReport.asp?inspecID=" & inspecID)
			End If
		End If ' On Error		
	End If ' Is Empty	
	Set upLoad = Nothing
End If ' Form Request %>
</body>
</html>