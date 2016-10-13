<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
%><!-- #include virtual="admin/connSWPPP.asp" --><%
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
baseDir = Request.ServerVariables("APPL_PHYSICAL_PATH") %>
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
	<!-- #include virtual="admin/adminHeader2.inc" -->

	<form method="post" action="<% = Request.ServerVariables("script_name") %>" onSubmit="return isReady(this)">
	<div class="six columns alpha">
		<h1>Add New Inspection Report</h1>
	</div>    
	<div class="six columns omega">
		<input name="submit" type="submit" value="Add New Report">
	</div>
	<div class="cleaner"></div>
	
	<div class="two columns alpha"><b>Date<small>(mm / dd / yyyy)</small>:</b></div>
	<div class="two columns"><input type="text" name="inspecDate" value="<%= Date()%>"></div>
	<div class="two columns"><b>Project Name:</b></div>
	<div class="six columns omega"><input type="text" name="projectName"></div>
	
	<div class="two columns alpha"><b>Project Phase:</b></div>
	<div class="four columns"><input type="text" name="projectPhase"></div>
	<div class="two columns"><b>Project Location:</b></div>
	<div class="four columns omega"><input type="text" name="projectAddr"></div>
		
	<div class="one columns alpha"><b>City:</b></div>
	<div class="two columns"><input type="text" name="projectCity"></div>
	<div class="one columns"><b>State:</b></div>
	<div class="two columns">
		<select name="projectState">
<% 	SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
	SET RS0=connSWPPP.execute(SQL0)
	DO WHILE NOT RS0.EOF %>
			<option value="<%= RS0("stateAbbr")%>"><%= RS0("stateAbbr")%></option>
<%	RS0.MoveNext
	LOOP %>	
		</select>
	</div>
	<div class="one columns"><b>Zip:</b></div>
	<div class="two columns"><input type="text" name="projectZip"></div>
	<div class="one columns"><b>County:</b></div>
	<div class="two columns omega">
		<select name="projectCounty">
<% 	SQL1="SELECT * FROM Counties WHERE stateAbbr='TX' OR stateAbbr='OK' ORDER BY priority DESC, countyName ASC"
	SET RS1=connSWPPP.execute(SQL1)
	DO WHILE NOT RS1.EOF %>
			<option value="<%= RS1("countyName")%>"><%= RS1("countyName")%></option>
<%	RS1.MoveNext
	LOOP %>	
		</select>
	</div>
	<div class="three columns alpha"><b>On-Site Contact:</b></div>
	<div class="nine columns omega"><input type="text" name="onsiteContact"></div>
	
	<div class="three columns alpha"><b>Contact Phone:</b></div>
	<div class="nine columns omega"><input name="officePhone" type="text"></div>
	
	<div class="three columns alpha"><b>Emergency Phone:</b></div>
	<div class="nine columns omega"><input name="emergencyPhone" type="text"></div>

	<h3>Company Information</h3>
	<div class="two columns alpha"><b>Company Name:</b></div>
	<div class="ten columns omega"><input type="text" name="compName"></div>

	<div class="two columns alpha"><b>Address 1:</b></div>
	<div class="ten columns omega"><input name="compAddr" type="text"></div>
	
	<div class="two columns alpha"><b>Address 2:</b></div>
	<div class="ten columns omega"><input name="compAddr2" type="text"></div>
	
	<div class="one columns alpha"><b>City:</b></div>
	<div class="three columns"><input name="compCity" type="text"></div>
	<div class="one columns"><b>State:</b></div>
	<div class="three columns">
		<select name="compState">
<% 	SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
	SET RS0=connSWPPP.execute(SQL0)
	DO WHILE NOT RS0.EOF %>
			<option value="<%= RS0("stateAbbr")%>"><%= RS0("stateAbbr")%></option>
<%	RS0.MoveNext
	LOOP %>					
		</select>
	</div>
	<div class="one columns"><b>Zip:</b></div>
	<div class="three columns omega"><input name="compZip" type="text"></div>
	<div class="cleaner"></div>
	
	<div class="two columns alpha"><b>Company Phone:</b></div>
	<div class="four columns"><input name="compPhone" type="text"></div>
	<div class="two columns"><b>Contact:</b></div>
	<div class="four columns omega"><input type="text" name="compContact"></div>
	
	<div class="two columns alpha"><b>Contact Phone:</b></div>
	<div class="four columns"><input name="contactPhone" type="text"></div>
	<div class="two columns"><b>Contact Fax:</b></div>
	<div class="four columns omega"><input name="contactFax" type="text"></div>
	
	<div class="two columns alpha"><b>Contact E-Mail:</b></div>
	<div class="ten columns omega"><input type="text" name="contactEmail"></div>

	<h3>Report Details</h3>
	<div class="two columns alpha"><b>Type of Report:</b></div>
	<div class="ten columns omega">
		<select name="reportType">
<% 	SQL2="SELECT * FROM ReportTypes ORDER BY priority DESC, reportTypeID ASC"
	SET RS2=connSWPPP.execute(SQL2)
	DO WHILE NOT RS2.EOF %>
			<option value="<%= RS2("reportType")%>"><%=RS2("reportType")%></option>
<% 	RS2.MoveNext
	LOOP %>
		</select>
	</div>
	
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
	<div class="two columns alpha"><b>Narrative</b></div>
	<div class="ten columns omega"><textarea rows="15" name="narrative"></textarea></div>
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
	<div class="two columns alpha"><b>Inspector:</b></div>
	<div class="ten columns omega">
		<select name="inspector">
		<% Do While Not connUser.EOF %>
			<option value="<% = connUser("userID") %>"> 
			<% = Trim(connUser("firstName")) & "&nbsp;" & Trim(connUser("lastName")) %>
			</option>
<%			connUser.moveNext
		Loop				
	connUser.Close
	Set connUser = Nothing %>
		</select>
	</div>
<%	End If
IF Session("validAdmin") OR Session("validInspector") THEN
	'Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	'Set objSiteMapDir = folderSvrObj.GetFolder(baseDir & "images\sitemap\")
	'Set siteMapImage = objSiteMapDir.Files 
	't1="sitemap"
	't2="sitemapDN"
	't3="sitemapUP" %>

<!-- ------------- Optional Links ----------------------------------------------------- -->
<% SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=0 ORDER BY oitSortByVal asc"
SET RS1=connSWPPP.execute(SQL1)%>

<h3>Optional Project Links</h3>
<table width="100%" border="0" align="center">
<%
DO WHILE NOT RS1.EOF
	oitID=RS1("oitID")
	dirName=Trim(RS1("oitName"))
	oitDesc=Trim(RS1("oitDesc"))
	t1=dirName
	t2=dirName &"DN"
	t3=dirName &"UP"

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objTemp = FSO.GetFolder(baseDir & "\images\" & dirName & "\")
	Set TempImage = objTemp.Files %>
	
    <tr valign="top">
		<td align="right" bgcolor="#eeeeee"><font size="-1"><strong><%= dirName%>:</strong></font></td>
		<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right">
			<BUTTON class="up-down-button" onClick="swapOption(<%= t1%>, <%= t2%>, 'up');">&uarr;</BUTTON><br>
			<BUTTON class="up-down-button" onClick="swapOption(<%= t1%>, <%= t2%>, 'dn');">&darr;</BUTTON>
		</td>	
		<td bgcolor="#999999" nowrap align="justify"><select name="<%=dirName%>DN"  size="3"></SELECT>
			<input type="hidden" name="<%= t1%>" value="">
		</td>
		<td bgcolor="#999999">
			<BUTTON class="left-right-button" onClick="delOption(<%= t1%>, <%= t2%>, <%= t3%>);">&rarr;</BUTTON>
			<BUTTON class="left-right-button" onClick="addOption(<%= t1%>, <%= t2%>, <%= t3%>);">&larr;</BUTTON>
		</td>
		<td bgcolor="#999999">
            <select name="<%=dirName%>UP">

<%	For Each Item In TempImage
		shortName = Item.Name %><option value="<% = shortName %>"><%= shortName %></option>
<%	Next 
	Set objTemp = Nothing
	Set TempImage = Nothing %>
		    </select>
		</td>
		<td bgcolor="#999999" nowrap align="right"><input type="button" style="width:120;"
				value="Upload" onClick="location = 'upImageRpt.asp?inspecID=<%= inspecID %>&oitID=<%= oitID %>'; return false";>
		</td>
	</tr>
    <% RS1.MoveNext
LOOP %>
	<tr> 
		<td colspan="5" align="center"><input name="submit" type="submit" value="Add New Report & Project Info"></td>
	</tr>
<% End If 'Session("validAdmin") %>
</table>
</form>
</body>
</html>
