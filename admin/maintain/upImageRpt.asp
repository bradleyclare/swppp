<%@ Language="VBScript" %>
<!-- #include virtual="admin/connSWPPP.asp" --> 
<!-- #include file="freeASPUpload.asp" --><%
  
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
	baseDir = Request.ServerVariables("APPL_PHYSICAL_PATH") %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : <%= imgType %> Upload</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
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
			<td width="75%"><input name="attach1" type="file" size="50"></td></tr>
		<tr><td height="44" align="left">&nbsp;</td>
			<td height="44" valign="bottom"> <input type="submit" value="Start Upload"></td></tr>
	</form>
</table>
<% Else 
		Server.ScriptTimeout=4500 'default is 90 seconds %>
<table width="100%" border="0">
<%	Set Upload = New FreeASPUpload
    Upload.Save( baseDir & imgDir )

	inspecID = UpLoad.Form("inspecID")	
	If inspecID="" Then
		Response.Redirect("newReport.asp")
	Else		
		Response.Redirect("editReport.asp?inspecID=" & inspecID)
	End If
	Set upLoad = Nothing
End If ' Form Request %>
</body>
</html>