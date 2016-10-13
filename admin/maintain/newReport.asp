<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
%><!-- #include file="../connSWPPP.asp" --><%
If Request.Form.Count > 0 Then
	Function strQuoteReplace(strValue)
		strQuoteReplace = Trim(Replace(strValue, "'", "''"))
	End Function
	projectName=strQuoteReplace(TRIM(Request("projectName")))
	projectPhase=strQuoteReplace(TRIM(Request("projectPhase")))

	userID = Session("userID")
	inspector = strQuoteReplace(Request("inspector"))

	If inspector <> "" Then userID = inspector End If

	' does project name already exist?
	SQLSELECT = "SELECT projectID FROM Projects" &_
		" WHERE projectName='" & strQuoteReplace(Request("projectName")) & "'"
	Response.Write(SQLSELECT & "<br>")
	SET connName=connSWPPP.execute(SQLSELECT)

	SQLSELECT = "SELECT projectID FROM Projects WHERE projectName='" & projectName &"' AND projectPhase='"& projectPhase &"'"
	Set connProjID=connSWPPP.execute(SQLSELECT)
	If not connProjID.eof then
		projectID=connProjID("projectID")
	else	
		SQLINSERT = "INSERT INTO Projects (" & _
			"projectName, projectPhase) VALUES (" & _
			"'" & projectName & "','" & projectPhase & "')"			
		' Response.Write(SQLINSERT & "<br><br>")
		connSWPPP.Execute(SQLINSERT)
		
		Set rsMaxProjectID = connSWPPP.Execute(SQLSELECT)
		projectID = rsMaxProjectID(0)
		
		rsMaxProjectID.Close
		Set rsMaxProjectID = Nothing
	end if 'set projectID
	connProjID.Close
	SET connProjID=nothing
	SQLa=""
	SQLb=""
	inspectSQLINSERT = "INSERT INTO Inspections (" & _
		"inspecDate, projectName, projectPhase, projectAddr, projectCity, projectState, projectZip, " & _
		"projectCounty, onsiteContact, officePhone, emergencyPhone, projectID, " & _
		"reportType, inches, bmpsInPlace, sediment, userID, compName, compAddr, " & _
		"compAddr2, compCity, compState, compZip, compPhone, compContact, " & _
		"contactPhone, contactFax, contactEmail" & _
		") VALUES (" & _
		"'" & strQuoteReplace(Request("inspecDate")) & "'" & _
		", '" & strQuoteReplace(Request("projectName")) & "'" & _
		", '" & strQuoteReplace(Request("projectPhase")) & "'" & _
		", '" & strQuoteReplace(Request("projectAddr")) & "'" & _
		", '" & strQuoteReplace(Request("projectCity")) & "'" & _
		", '" & Request("projectState") & "'" & _
		", '" & strQuoteReplace(Request("projectZip")) & "'" & _
		", '" & Request("projectCounty") & "'" & _
		", '" & strQuoteReplace(Request("onsiteContact")) & "'" & _
		", '" & strQuoteReplace(Request("officePhone")) & "'" & _
		", '" & strQuoteReplace(Request("emergencyPhone")) & "'" & _
		", " & projectID & _
		", '" & Request("reportType") & "'" & _
		", '-1'" & _
		", '-1'" & _
		", '-1'" & _
		", " & userID & _
		", '" & strQuoteReplace(Request("compName")) & "'" & _
		", '" & strQuoteReplace(Request("compAddr")) & "'" & _
		", '" & strQuoteReplace(Request("compAddr2")) & "'" & _
		", '" & strQuoteReplace(Request("compCity")) & "'" & _
		", '" & Request("compState") & "'" & _
		", '" & strQuoteReplace(Request("compZip")) & "'" & _
		", '" & strQuoteReplace(Request("compPhone")) & "'" & _
		", '" & strQuoteReplace(Request("compContact")) & "'" & _
		", '" & strQuoteReplace(Request("contactPhone")) & "'" & _
		", '" & strQuoteReplace(Request("contactFax")) & "'" & _
		", '" & strQuoteReplace(Request("contactEmail")) & "'" & _
		")"
		
	Response.Write(inspectSQLINSERT & "<br><br>")
	connSWPPP.Execute(inspectSQLINSERT)

		maxInspectSQLSELECT = "SELECT MAX(inspecID) FROM Inspections"
		Set rsMaxInspectID = connSWPPP.Execute(maxInspectSQLSELECT)
		maxInspectID = rsMaxInspectID(0)
	SQL1="UPDATE Inspections SET narrative='"& REPLACE(Request("narrative"),"'","#@#") &"'" &_
		" WHERE inspecID='"& maxInspectID &"'"
	connSWPPP.execute(SQL1)

'--	Creat Optional Image Associations -------------------------------
		SQLa="SELECT * FROM OptionalImagesTypes ORDER BY oitSortByVal asc"
		SET RSa=connSWPPP.Execute(SQLa)
		DO WHILE NOT RSa.EOF
			cnt=1
			tList=SPLIT(Request(Trim(RSa("oitName"))),":")
			For m=1 to Ubound(tList) Step 1
'--	dbo.sp_AddOptImage 	@oImageName char(50),@oIamgeDesc char(50),@inspecID int,@oitID int,@oImageFileName char(50),@oOrder smallint
				SQLb=SQLb &"/*<br>*/ EXEC sp_AddOptImage null,null,"& maxInspectID &","& RSa("oitID") &",'"& Trim(tList(m-1)) &"',"& cnt
				cnt=cnt+1
			Next
			RSa.MoveNext
		LOOP 
Response.Write(SQLb)
		IF LEN(SQLb)>1 THEN connSWPPP.execute(SQLb) END IF
'-- --------- Check for an existing project user inspector record match ----------------------------	
	SQL1="SELECT * FROM ProjectsUsers WHERE userID='"& userID &"' AND "&_
		"projectID='"& projectID &"' AND rights='inspector'"
	SET RS1=connSWPPP.Execute(SQL1)
'-- --------- If no match then add ProjectUsers record ---------------------------------------------
	IF (RS1.BOF AND RS1.EOF) THEN
	coUserSQLINSERT = "INSERT INTO ProjectsUsers (" & _
		"userID, projectID, rights" & _
		") VALUES (" & _
		userID & _
		", " & projectID & _
		", 'inspector')"
	Response.Write(coUserSQLINSERT & "<br><br>")
	connSWPPP.Execute(coUserSQLINSERT)
	END IF '--- ------ ProjectsUsers record already exists -----------------------------------------
	
	Session("inspecID")=maxInspectID
	rsMaxInspectID.Close
	Set rsMaxInspectID = Nothing
	
	Response.Redirect("addCoordinate.asp?inspecID=" & maxInspectID)
End If
'baseDir = "c:\Inetpub\wwwroot\SWPPP\"
baseDir = "d:\inetpub\wwwroot\swppp\" %>
<html>
<head>
<title>SWPPP INSPECTIONS : Add New Inspection Report</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
<script language="JavaScript" src="../js/validReports.js"></script>
<script language="JavaScript" src="../js/validReports1.2.js"></script>
<script language="JavaScript1.2"><!--
// we Can't just use the same transfer functio for both directions because
// the hidden input keys off of the t2 value solely...
function addOption(t1, t2, t3) {
    var index = t3.selectedIndex;
    if (index > -1) {
        var newoption = new Option(t3.options[index].text, t3.options[index].value, true, true);
        t2.options[t2.length] = newoption;
        if (!document.getElementById) history.go(0);
        t3.options[index] = null;
        t3.selectedIndex = 0;
		var tempStr="";
		for(var i=0; i<(t2.length) ;i++){
			tempStr=tempStr + (t2.options[i].value) + ":" ;
		}		
		t1.value=tempStr;
    }
}function delOption(t1, t3, t2) {
    var index = t3.selectedIndex;
    if (index > -1) {
        var newoption = new Option(t3.options[index].text, t3.options[index].value, true, true);
        t2.options[t2.length] = newoption;
        if (!document.getElementById) history.go(0);
        t3.options[index] = null;
        t3.selectedIndex = 0;
		var tempStr="";
		for(var i=0; i<(t3.length) ;i++){
			tempStr=tempStr + (t3.options[i].value) + ":" ;
		}
		t1.value=tempStr;
    }
}
function swapOption(t1, t2, slideDir) {
	var curIndex = t2.selectedIndex;
	var swapIndex= curIndex;
	var maxIndex= t2.length;
	if (curIndex > -1) {
		(slideDir=="up") ? (swapIndex=curIndex-1):(swapIndex=curIndex+1);
		if ((swapIndex>-1) && (swapIndex<t2.length)) {
			var newOption = new Option(t2.options[swapIndex].text, t2.options[swapIndex].value, true, true);
			t2.options[maxIndex] = newOption;
			t2.options[swapIndex].text=t2.options[curIndex].text;
			t2.options[swapIndex].value=t2.options[curIndex].value;
			t2.options[curIndex].text=t2.options[maxIndex].text;
			t2.options[curIndex].value=t2.options[maxIndex].value;
			t2.options[maxIndex] = null;
			t2.selectedIndex=swapIndex;
			var tempStr="";
			for(var i=0; i<(t2.length) ;i++){
				tempStr=tempStr + (t2.options[i].value) + ":" ;
			}
			t1.value=tempStr;
		}
	}	
}
//--></script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Add New Inspection Report</h1>
    
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
	<form method="post" action="<% = Request.ServerVariables("script_name") %>" onSubmit="return isReady(this)";>	 
		<tr><td colspan="3" align="center">
			<input name="submit" type="submit" value="Add New Report"><br><br></td></tr>
		<!-- date -->
		<tr> 
			<td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Date:</b></td>
			<td bgcolor="#999999"> <input type="text" name="inspecDate" size="10" maxlength="10" value="<%= Date()%>"> 
				<small>&nbsp;(mm / dd / yyyy)</small></td>
		</tr>
		<!-- project name -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Project Name | Phase:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectName" size="50" maxlength="50">
								<input type="text" name="projectPhase" size="20" maxlength="20"></td>
		</tr>
		<!-- project location -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Project Location:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectAddr" size="50" maxlength="50"> 
			</td>
		</tr>
		<tr>
			<td align="right" bgcolor="#eeeeee"><b>City, State, Zip:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectCity" size="20" maxlength="20">
			&nbsp; <select name="projectState">
<% 	SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
	SET RS0=connSWPPP.execute(SQL0)
	DO WHILE NOT RS0.EOF %>
		<option value="<%= RS0("stateAbbr")%>"><%= RS0("stateAbbr")%></option>
<%	RS0.MoveNext
	LOOP %>	
			</select> &nbsp; <input type="text" name="projectZip" size="5" maxlength="5"> 
			</td>
		</tr>
		<!-- onsite contact -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>County:</b></td>
			<td bgcolor="#999999"><select name="projectCounty">
<% 	SQL1="SELECT * FROM Counties WHERE stateAbbr='TX' OR stateAbbr='OK' ORDER BY priority DESC, countyName ASC"
	SET RS1=connSWPPP.execute(SQL1)
	DO WHILE NOT RS1.EOF %>
		<option value="<%= RS1("countyName")%>"><%= RS1("countyName")%></option>
<%	RS1.MoveNext
	LOOP %>	
			</select></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="onsiteContact" size="50" maxlength="50"></td>
		</tr>
		<!-- office # -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input name="officePhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- emergency # -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"> <b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input name="emergencyPhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- company -->
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"><font size="+1">Company Information</font></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Company Name:</b></td>
			<td bgcolor="#999999"><input type="text" name="compName" size="50" maxlength="50"></td>
		</tr>
		<!-- Address -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Address 1:</b></td>
			<td bgcolor="#999999"><input name="compAddr" type="text" size="50" maxlength="50"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Address 2:</b></td>
			<td bgcolor="#999999"><input name="compAddr2" type="text" size="50" maxlength="50"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>City:</b></td>
			<td bgcolor="#999999"><input name="compCity" type="text" size="20" maxlength="20"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>State:</b></td>
			<td bgcolor="#999999"><select name="compState">
<% 	SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
	SET RS0=connSWPPP.execute(SQL0)
	DO WHILE NOT RS0.EOF %>
		<option value="<%= RS0("stateAbbr")%>"><%= RS0("stateAbbr")%></option>
<%	RS0.MoveNext
	LOOP %>					</select></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Zip:</b></td>
			<td bgcolor="#999999"><input name="compZip" type="text" size="5" maxlength="5"></td>
		</tr>
		<!-- main telephone number -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Company Phone:</b></td>
			<td bgcolor="#999999"><input name="compPhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- contact -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="compContact" size="50" maxlength="50"></td>
		</tr>
		<!-- phone -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact Phone:</b></td>
			<td bgcolor="#999999"><input name="contactPhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- fax -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact Fax:</b></td>
			<td bgcolor="#999999"><input name="contactFax" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- e-mail -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact E-Mail:</b></td>
			<td bgcolor="#999999"><input type="text" name="contactEmail" size="50" maxlength="50"></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
</Table>
<!-- Type of Report? Weekly, Storm, Complaint, ? -->
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">		
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999" colspan=2><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Type of Report:</b></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right"></td>
			<td bgcolor="#999999"><select name="reportType">
<% 	SQL2="SELECT * FROM ReportTypes ORDER BY priority DESC, reportTypeID ASC"
	SET RS2=connSWPPP.execute(SQL2)
	DO WHILE NOT RS2.EOF %>
			<option value="<%= RS2("reportType")%>"><%=RS2("reportType")%></option>
<% 	RS2.MoveNext
	LOOP %>
			</select></td>
		</tr>
		<!-- Rain -->
<!--		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Inches of Rain:</b></td>
			<td bgcolor="#999999"><input type="text" name="inches" size="6" maxlength="6" value="0"></td>
		</tr>-->
		<!-- BMPs? y/n -->
<!--		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Are BMPs in place?</b></td>
			<td bgcolor="#999999"><select name="bmpsInPlace">
					<option value="1">Yes</option>
					<option value="0">No</option>
				</select></td>
		</tr>-->
		<!-- sediment loss or pollution? y/n -->
<!--		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Sediment Loss or Pollution?</b></td>
			<td bgcolor="#999999"><select name="sediment">
					<option value="1">Yes</option>
					<option value="0">No</option>
				</select></td>
		</tr>-->
		<TR><TD align="right" bgcolor="#eeeeee"><b>Narrative</b></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right"></td>
			<td bgcolor="#999999">
				<textarea cols="60" rows="5" name="narrative"></textarea><br><br>
		</TD></TR>
<%
' Admin can change inspector name.
If Session("validAdmin") Then
	
	insSQLSELECT = "SELECT DISTINCT u.userID," & _
		" firstName, lastName" & _
		" FROM Users as u, ProjectsUsers as pu" & _
		" WHERE u.userID = pu.userID" & _
		" AND pu.rights = 'inspector'"	
	Set connUser = connSWPPP.execute(insSQLSELECT)
%>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>Inspector:</strong></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right"></td>
			<td bgcolor="#999999"><select name="inspector">
					<% Do While Not connUser.EOF %>
					<option value="<% = connUser("userID") %>"> 
						<% = Trim(connUser("firstName")) & "&nbsp;" & Trim(connUser("lastName")) %>
					</option>
<%				 	connUser.moveNext
				Loop				
	connUser.Close
	Set connUser = Nothing %>
				</select></td></tr>
<%	End If
	IF Session("validAdmin") OR Session("validInspector") THEN
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	Set objSiteMapDir = folderSvrObj.GetFolder(baseDir & "images\sitemap\")
	Set siteMapImage = objSiteMapDir.Files 
	t1="sitemap"
	t2="sitemapDN"
	t3="sitemapUP" %>
<!--		<tr valign="top"><td align="right" bgcolor="#eeeeee"><strong>Site Map File:</strong></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right" valign="top">
				<BUTTON onClick="swapOption(<%= t1%>, <%= t2%>, 'up');">&uarr;</BUTTON><br>
				<BUTTON onClick="swapOption(<%= t1%>, <%= t2%>, 'dn');">&darr;</BUTTON></td>
			<td bgcolor="#999999" nowrap valign="top">
				<select name="sitemapDN" size="2"></SELECT>
				<input type="hidden" name="sitemap" value="">
					<BUTTON onClick="delOption(<%= t1%>, <%= t2%>, <%= t3%>);">--&gt;</BUTTON>
					<BUTTON onClick="addOption(<%= t1%>, <%= t2%>, <%= t3%>);">&lt;--</BUTTON>
				<select name="sitemapUP">
<%	For Each Item In siteMapImage
		shortName = Item.Name 
		IF InStr(tempStrOfFileNames, shortName)=0 THEN %>
				<option value="<% = shortName %>"><% = shortName %></option>
<%		END IF
	Next
	Set objSteMapDir = Nothing
	Set siteMapImage = Nothing %>
				</select>&nbsp;&nbsp;<input type="button" value="Upload Site Map File" 
					onClick="location='upSiteMapEditRprt.asp?inspecID=<% = inspecID %>'; return false";>
				</td></tr>-->
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999" colspan=2><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr><td colspan="3"><hr align="center" width="95%" size="1"></td></tr>
		<tr><td colspan="3"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<!-- Type of Report? Weekly, Storm, Complaint, ? -->
</Table>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">
		<tr><td colspan="6" align="center"><font size="+1">Optional Project Links</font></td>
		</tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999" colspan=5><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
<!-- ------------- Optional Links ----------------------------------------------------- -->

<%	
SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=0 ORDER BY oitSortByVal asc"
SET RS1=connSWPPP.execute(SQL1)
DO WHILE NOT RS1.EOF
	oitID=RS1("oitID")
	dirName=Trim(RS1("oitName"))
	oitDesc=Trim(RS1("oitDesc"))
	t1=dirName
	t2=dirName &"DN"
	t3=dirName &"UP"

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objTemp = FSO.GetFolder(baseDir & "images\"& dirName &"\")
	Set TempImage = objTemp.Files %>
		<tr valign="top">
			<td align="right" bgcolor="#eeeeee"><font size="-1"><strong><%= dirName%>:</strong></font></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right">
				<BUTTON onClick="swapOption(<%= t1%>, <%= t2%>, 'up');">&uarr;</BUTTON><br>
				<BUTTON onClick="swapOption(<%= t1%>, <%= t2%>, 'dn');">&darr;</BUTTON></td>	
			<td bgcolor="#999999" nowrap align="justify">
				<select name="<%=dirName%>DN"  size="3"></SELECT>
					<input type="hidden" name="<%= t1%>" value=""></td>
					<td bgcolor="#999999">
						<BUTTON onClick="delOption(<%= t1%>, <%= t2%>, <%= t3%>);">--&gt;</BUTTON>
						<BUTTON onClick="addOption(<%= t1%>, <%= t2%>, <%= t3%>);">&lt;--</BUTTON></td>
					<td bgcolor="#999999"><select name="<%=dirName%>UP">
<%	For Each Item In TempImage
		shortName = Item.Name %><option value="<% = shortName %>"><%= shortName %></option>
<%	Next 
	Set objTemp = Nothing
	Set TempImage = Nothing %>
				</select></td>
			<td bgcolor="#999999" nowrap align="right"><input type="button" style="width:120;"
				value="Upload File/Image" onClick="location='upImageRpt.asp?inspecID=<%= inspecID %>&oitID=<%= oitID %>'; return false";>
				</td></tr><% 
	RS1.MoveNext
LOOP
	End If 'Session("validAdmin") %>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999" colspan=5><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr> 
			<td colspan="5">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="5" align="center">
			<input name="submit" type="submit" value="Add New Report & Project Info"></td>
		</tr>
	</form>
</table>
</body>
</html>
