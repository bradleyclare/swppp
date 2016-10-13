<!-- #include virtual="admin/connSWPPP.asp" --><%
	inspecID=4003
	inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, projectAddr, projectCity, projectState" & _
		", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone" & _
		", i.projectID, compName, compAddr, compAddr2, compCity, compState, compZip, compPhone" & _
		", compContact, contactPhone, contactFax, contactEmail, reportType, inches, bmpsInPlace" & _
		", sediment, userID" & _
		" FROM Inspections as i, Projects as p" & _
		" WHERE i.projectID = p.projectID" & _
		" AND inspecID = " & inspecID
'--Response.Write(inspecSQLSELECT & "<br>")
	Set rsReport = connSWPPP.execute(inspecSQLSELECT) %>
<html>
<head>
	<title>SWPPP INSPECTIONS : Edit Inspection Report</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validReports.js"></script>
	<script language="JavaScript" src="../js/validReports1.2.js"></script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Edit Inspection Report</h1>	
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
	<form method="post" action="<% = Request.ServerVariables("script_name") %>" 
	onSubmit="return isReady(this)";>
		<input type="hidden" name="inspecID" value="<% = inspecID %>">
		<input type="hidden" name="projectID" value="<% = rsReport("projectID") %>">
<%	IF Session("validAdmin") OR Session("validInspector") THEN
	tempDev = "dev\"
	baseDir = Request.ServerVariables("APPL_PHYSICAL_PATH") &tempDev
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	Set objSteMapDir = folderSvrObj.GetFolder(baseDir & "images\sitemaps\")
	Set siteMapImage = objSteMapDir.Files 

	SQLa="sp_oImagesByType ("& inspecID &",'sitemap')" 
	SET RSa=connSWPPP.execute(SQLa) %>
		<tr><td align="right" bgcolor="#eeeeee"><strong>Site Map File:</strong></td>
			<td bgcolor="#999999" nowrap>
				<SPAN id="sitemapSPAN">
				<select name="siteMap" size="1" multiple>
<% 	DO WHILE NOT(RSa.EOF) %>
					<OPTION value="<%= RSa("oImageFileName")%>"><%= RSa("oImageFileName")%></OPTION>
<%		RSa.MoveNext
	LOOP %>
				</SELECT><BUTTON>delete</BUTTON><BUTTON>add</BUTTON><select name="siteMapUP" size="1" multiple>
<%	For Each Item In siteMapImage
		shortName = Item.Name %>
				<option value="<% = shortName %>"><% = shortName %></option>
<%	Next
	Set objSteMapDir = Nothing
	Set siteMapImage = Nothing %>
				</select></SPAN> &nbsp;&nbsp; <input type="button" value="Upload Site Map File" 
					onClick="location='upSiteMapEditRprt.asp?inspecID=<% = inspecID %>'; return false";>
				</td></tr>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><!-- Type of Report? Weekly, Storm, Complaint, ? --><tr> 
		<tr><td colspan="2" align="center"><font size="+1">Optional Project Links</font></td>
		</tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
<!-- ------------- Optional Links ----------------------------------------------------- -->
 
<%	
DIM tempArray(7)
tempArray(0)="SWPPP"
tempArray(1)="NOI"
tempArray(2)="Construction Site Notice"
tempArray(3)="Permit"
tempArray(4)="Construction Sign"
tempArray(5)="Subcontractor Certification"
tempArray(6)="NOT"
tempArray(7)="Operator Form"
FOR n= 0 to 7 step 1	
	tempTrimmed= REPLACE(tempArray(n)," ","")
	tempDev = "dev\"
	baseDir = Request.ServerVariables("APPL_PHYSICAL_PATH") & tempDev
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
'Response.Write(baseDir & "images\"& tempTrimmed &"\")
	Set objTemp = FSO.GetFolder(baseDir & "images\"& tempTrimmed &"\")
	Set TempImage = objTemp.Files 
'--	SQLa="sp_oImagesByType "& inspecID &",'"& tempTrimmed &"'" 
'--Response.Write(SQLa &"<br>")
'--	SET RSa=connSWPPP.execute(SQLa)	
%>
		<tr valign="middle"><td align="right" bgcolor="#eeeeee"><strong><%= tempArray(n)%>:</strong></td>
			<td bgcolor="#999999" nowrap align="justify">
				<select name="<%=tempTrimmed%>"  size="1" multiple>
				<SPAN id="<%=tempTrimmed%>SPAN">
<% tempStr="z2.asp" %>
<!-- #include file= -->
				</SPAN></select>&nbsp;&nbsp;<input type="button" style="width:120;"
					value="Upload New Image File" onClick="location='upImageRpt.asp?inspecID=<%= inspecID %>&imageTypeID=<%= n %>'; return false";>
				</td></tr>
<% NEXT
	End If 'Session("validAdmin") %>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr><td colspan="2">&nbsp;</td>
		</tr><tr><td>&nbsp;</td>
			<td><input name="submit" type="submit" value="Edit Report & Project Info"></td>
		</tr>
	</form>
</table>
</body>
</html>