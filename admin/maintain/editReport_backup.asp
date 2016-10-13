<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID = Session("inspecID")
IF Request("inspecID")<>"" THEN 
	inspecID = Request("inspecID") 
	Session("inspecID")=inspecID
END IF
%><!-- #include virtual="admin/connSWPPP.asp" --><%
If Request.Form.Count > 0 Then	
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	
	userID = Session("userID")
	inspector = TRIM(strQuoteReplace(Request("inspector")))	
	If inspector <> "" Then userID = inspector End If
	bmps=Request("bmpsInPlace")
	sediment=Request("sediment")
	upProjPhase= strQuoteReplace(Request("projectPhase"))	
	inspectSQLUPDATE = "UPDATE Inspections SET" & _
		" inspecDate = '" & strQuoteReplace(Request("inspecDate")) & "'" & _
		", projectName = '" & strQuoteReplace(Request("projectName")) & "'" 
		IF LEN(TRIM(upProjPhase))=0 THEN 
			inspectSQLUPDATE=inspectSQLUPDATE &", projectPhase=null" 
		ELSE
			inspectSQLUPDATE=inspectSQLUPDATE &", projectPhase='" & upProjPhase &"'"
		END IF
		inspectSQLUPDATE=inspectSQLUPDATE &", projectAddr = '" & strQuoteReplace(Request("projectAddr")) & "'" & _
		", projectCity = '" & strQuoteReplace(Request("projectCity")) & "'" & _
		", projectState = '" & Request("projectState") & "'" & _
		", projectZip = '" & strQuoteReplace(Request("projectZip")) & "'" & _
		", projectCounty = '" & Request("projectCounty") & "'" & _
		", onsiteContact = '" & strQuoteReplace(Request("onsiteContact")) & "'" & _
		", officePhone = '" & strQuoteReplace(Request("officePhone")) & "'" & _
		", emergencyPhone = '" & strQuoteReplace(Request("emergencyPhone")) & "'" & _
		", reportType = '" & Request("reportType") & "'" & _
		", inches = " & Request("inches") & _
		", bmpsInPlace = " & bmps & _
		", sediment = " & sediment & _
		", userID = " & inspector & _
		", compName = '" & strQuoteReplace(Request("compName")) & "'" & _
		", compAddr = '" & strQuoteReplace(Request("compAddr")) & "'" & _
		", compAddr2 = '" & strQuoteReplace(Request("compAddr2")) & "'" & _
		", compCity = '" & strQuoteReplace(Request("compCity")) & "'" & _
		", compState = '" & Request("compState") & "'" & _
		", compZip = '" & strQuoteReplace(Request("compZip")) & "'" & _
		", compPhone = '" & strQuoteReplace(Request("compPhone")) & "'" & _
		", compContact = '" & strQuoteReplace(Request("compContact")) & "'" & _
		", contactPhone = '" & strQuoteReplace(Request("contactPhone")) & "'" & _
		", contactFax = '" & strQuoteReplace(Request("contactFax")) & "'" & _
		", contactEmail = '" & strQuoteReplace(Request("contactEmail")) & "'" & _
		" WHERE inspecID = " & inspecID
'response.Write(inspectSQLUPDATE)
	connSWPPP.Execute(inspectSQLUPDATE)
'--	Creat Optional Image Associations -------------------------------
		SQLa="DELETE FROM OptionalImages WHERE inspecID="& inspecID 
		connSWPPP.Execute(SQLa)
		SQLa="SELECT * FROM OptionalImagesTypes ORDER BY oitSortByVal asc"
		SET RSa=connSWPPP.Execute(SQLa)
		DO WHILE NOT RSa.EOF
			cnt=1
'Response.Write(Request(Trim(RSa("oitName"))) & "<BR/>")
			tList=SPLIT(Request(Trim(RSa("oitName"))),":")
			For m=1 to Ubound(tList) Step 1
'--	dbo.sp_AddOptImage 	@oImageName char(50),@oIamgeDesc char(50),@inspecID int,@oitID int,@oImageFileName char(50),@oOrder smallint
				SQLb=SQLb &"/*<br>*/ EXEC sp_AddOptImage null,null,"& inspecID &","& RSa("oitID") &",'"& Trim(tList(m-1)) &"',"& cnt
				cnt=cnt+1
			Next
			RSa.MoveNext
		LOOP 
'Response.Write(SQLb)
        if Len(SQLb) > 1 then connSWPPP.execute(SQLb) End If
    
		for n = 0 to 99 step 1
'Response.Write("coord:coID:" & CStr(n)&":"& Request("coord:coID:" & CStr(n)) &"<br/>")
		    if Trim(Request("coord:coID:" & CStr(n))) = "" then
		        exit for
		    end if
'-- dbo.spAEDCoordinate @_iCOID int, @_DelFlag smallint, @_inspecID int, @_iCoordinates char(50), @_icorrectiveMods char(255), @_iOrderBy int
            DelCoord = 0
            if Request("coord:del:"& CStr(n)) = "on" then DelCoord = 1 End If
            OrderBy = 0
            if IsNumeric(Request("coord:orderby:"& CStr(n))) then OrderBy = Request("coord:orderby:"& CStr(n)) End If
		    SQLc = SQLc &"/*<br/>*/Exec spAEDCoordinate "& Request("coord:coID:"& CStr(n)) &", "& DelCoord &", "& inspecID &", '"& Replace(Request("coord:coord:"& CStr(n)),"--","—") &"', '"& Replace(Request("coord:mods:"& CStr(n)),"--","—") &"', "& OrderBy &";"
		next	
'Response.Write(SQLc)
        if Len(SQLc) > 0 then connSWPPP.execute(SQLc) end if
'Response.End	
    If request("submit") = "Edit Report & Project Info" Then	
	    connSWPPP.Close
	    Set connSWPPP = Nothing
    	Response.Redirect("viewReports.asp")
    End If
End If
	inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, projectAddr, projectCity, projectState" & _
		", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, i.projectID, compName" & _
		", compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax" & _
		", contactEmail, reportType, inches, bmpsInPlace, sediment, userID" & _
		" FROM Inspections as i, Projects as p" & _
		" WHERE i.projectID = p.projectID AND inspecID = " & inspecID
'--Response.Write(inspecSQLSELECT & "<br>")
	Set rsReport = connSWPPP.execute(inspecSQLSELECT)
'baseDir = "d:\vol\swpppinspections.com\www\htdocs\" 
baseDir = "D:\Inetpub\wwwroot\SWPPP\"%>
<html>
<head>
	<title>SWPPP INSPECTIONS : Edit Inspection Report</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<STYLE>
	select.long	{ font-size:xx-small;	}
	</STYLE>
	<script type="text/javascript" language="JavaScript" src="../js/validReports.js"></script>
	<script type="text/javascript" language="JavaScript" src="../js/validReports1.2.js"></script>
	<script type="text/javascript" language="JavaScript1.2"><!--
// we Can't just use the same transfer function for both directions because
// the hidden input keys off of the t2 value solely...-->
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
function editNarrative(inspID){
var basePath = "http://www.swppp.com";
var URL = "/admin/maintain/editNarrative.asp?inspecID=" + inspID;
var params = "height=420,width=520,status=yes,toolbar=no,menubar=no, directories=no, location=no, scrollbars=no, resizable=no";
	window.open(URL, "", params);
}
</script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Edit Inspection Report</h1>	
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
	<form id="theForm" method="post" action="<% = Request.ServerVariables("script_name") %>" onsubmit="return isReady(this)";>
		<input type="hidden" name="inspecID" value="<% = inspecID %>"/>
		<input type="hidden" name="projectID" value="<% = rsReport("projectID") %>"/>
		<tr><td></td><td><input type="button" value="Delete Report" onclick="location='deleteReport.asp?inspecID=<% = inspecID %>'; return false;"/><br/></td>
		</tr>
		<!-- date -->
		<tr><td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Date:</b></td>
			<td bgcolor="#999999"> <input type="text" name="inspecDate" size="10" value="<% = Trim(rsReport("inspecDate")) %>"> <small>&nbsp;(mm / dd / yyyy)</small></td>
		</tr>
		<!-- project name -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Project Name | Phase:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectName" size="50" value="<% = Trim(rsReport("projectName")) %>"/>
			<input type="text" name="projectPhase" size="20" value="<% = Trim(rsReport("projectPhase")) %>"/></td>
		</tr>
		<!-- project location -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Project Location:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectAddr" size="50" value="<% = Trim(rsReport("projectAddr")) %>"> </td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>City, State, Zip:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectCity" size="20" value="<% = Trim(rsReport("projectCity")) %>"> &nbsp; <select name="projectState">
<% 	SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
	SET RS0=connSWPPP.execute(SQL0)
	IF IsNull(TRIM(rsReport("projectState"))) THEN rsReport("projectState")="TX" END IF
	DO WHILE NOT RS0.EOF %>	<option value="<%= RS0("stateAbbr")%>"<% 
		If Trim(rsReport("projectState")) = RS0("stateAbbr") Then %> selected<% 
		End If %>><%= Trim(RS0("stateAbbr"))%></option>
<%	RS0.MoveNext
	LOOP %>	</select> &nbsp; <input type="text" name="projectZip" size="5" value="<% = Trim(rsReport("projectZip")) %>"> </td>
		</tr>
		<!-- onsite contact -->
		<tr><td align="right" bgcolor="#eeeeee"><b>County:</b></td>
			<td bgcolor="#999999"><select name="projectCounty">
<% 	SQL1="SELECT * FROM Counties ORDER BY priority DESC, countyName ASC"
	SET RS1=connSWPPP.execute(SQL1)
	DO WHILE NOT RS1.EOF %><option value="<%= Trim(RS1("countyName"))%>"<% 
	If Trim(rsReport("projectCounty")) = TRIM(RS1("countyName")) Then %> selected<% 
	End If %>><%= Trim(RS1("countyName"))%></option>
<%	RS1.MoveNext
	LOOP %>	</select></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="onsiteContact" size="50" value="<% = Trim(rsReport("onsiteContact")) %>"></td>
		</tr>
		<!-- office # -->
		<tr><td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input name="officePhone" type="text" size="50" value="<% = Trim(rsReport("officePhone")) %>"></td>
		</tr>
		<!-- emergency # -->
		<tr><td align="right" bgcolor="#eeeeee"> <b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input name="emergencyPhone" type="text" size="50" value="<% = Trim(rsReport("emergencyPhone")) %>"></td>
		</tr><tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<!-- company -->
		<tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr><tr><td colspan="2" align="center"><font size="+1">Company Information</font></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Company Name:</b></td>
			<td bgcolor="#999999"><input type="text" name="compName" size="50" value="<% = Trim(rsReport("compName")) %>"></td>
		</tr>
		<!-- Address -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Address 1:</b></td>
			<td bgcolor="#999999"><input name="compAddr" type="text" size="50" value="<% = Trim(rsReport("compAddr")) %>"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Address 2:</b></td>
			<td bgcolor="#999999"><input name="compAddr2" type="text" size="50" value="<% = Trim(rsReport("compAddr2")) %>"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>City:</b></td>
			<td bgcolor="#999999"><input name="compCity" type="text" size="20" value="<% = Trim(rsReport("compCity")) %>"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>State:</b></td>
			<td bgcolor="#999999"><select name="compState">
<% 	SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
	SET RS0=connSWPPP.execute(SQL0)
	IF IsNull(TRIM(rsReport("compState"))) THEN rsReport("compState")="TX" END IF
	DO WHILE NOT RS0.EOF %><option value="<%= Trim(RS0("stateAbbr"))%>"<% 
	If Trim(rsReport("compState")) = RS0("stateAbbr") Then %> selected<% 
	End If %>><%= Trim(RS0("stateAbbr"))%></option>
<%	RS0.MoveNext
	LOOP %>	</select></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Zip:</b></td>
			<td bgcolor="#999999"><input name="compZip" type="text" size="5" value="<% = Trim(rsReport("compZip")) %>"></td>
		</tr>
		<!-- main telephone number -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Company Phone:</b></td>
			<td bgcolor="#999999"><input name="compPhone" type="text" size="20" value="<% = Trim(rsReport("compPhone")) %>"></td>
		</tr>
		<!-- contact -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="compContact" size="50" value="<% = Trim(rsReport("compContact")) %>"></td>
		</tr>
		<!-- phone -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact Phone:</b></td>
			<td bgcolor="#999999"><input name="contactPhone" type="text" size="20" value="<% = Trim(rsReport("contactPhone")) %>"></td>
		</tr>
		<!-- fax -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact Fax:</b></td>
			<td bgcolor="#999999"><input name="contactFax" type="text" size="20" value="<% = Trim(rsReport("contactFax")) %>"></td>
		</tr>
		<!-- e-mail -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact E-Mail:</b></td>
			<td bgcolor="#999999"><input type="text" name="contactEmail" size="50" value="<% = Trim(rsReport("contactEmail")) %>"></td>
		</tr><tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<!-- Type of Report? Weekly, Storm, Complaint, ? -->
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Type of Report:</b></td>
			<td bgcolor="#999999"><select name="reportType">
<% 	SQL2="SELECT * FROM ReportTypes where priority > 0 ORDER BY priority DESC, reportTypeID ASC"
	SET RS2=connSWPPP.execute(SQL2)
	DO WHILE NOT RS2.EOF %><option value="<%= Trim(RS2("reportType"))%>"<% 
	If Trim(rsReport("reportType")) = TRIM(RS2("reportType")) Then %> selected<% End If %>><%= Trim(RS2("reportType"))%></option>
<% 	RS2.MoveNext
	LOOP %>	</select></td>
		</tr>
		<!-- Rain -->
<% IF Trim(rsReport("inches")) > "-1" THEN %>
		<tr><td align="right" bgcolor="#eeeeee"><b>Inches of Rain:</b></td>
			<td bgcolor="#999999"><input type="text" name="inches" size="6"	value="<% = Trim(rsReport("inches")) %>"></td></tr>
<% ELSE %>
		<INPUT type="hidden" name="inches" value="<%= Trim(rsReport("inches"))%>">
<% END IF %>
		<!-- BMPs? y/n -->
<% IF rsReport("bmpsInPlace") = "-1" THEN %>
		<INPUT type="hidden" name="bmpsInPlace" value="<%= rsReport("bmpsInPlace")%>">
<% ELSE %>
		<tr><td align="right" bgcolor="#eeeeee"><b>Are BMPs in place?</b></td>
			<td bgcolor="#999999"> <select name="bmpsInPlace">
					<option value="1"<% If rsReport("bmpsInPlace")="1" Then %> selected<% End If %>>Yes</option>
					<option value="0"<% If rsReport("bmpsInPlace")="0" Then %> selected<% End If %>>No</option>
				</select></td></tr>
<% END IF %>
		<!-- sediment loss or pollution? y/n -->
<% IF rsReport("bmpsInPlace") = "-1" THEN %>
		<INPUT type="hidden" name="sediment" value="<%= rsReport("sediment")%>">
<% ELSE %>
		<tr><td align="right" bgcolor="#eeeeee"><b>Sediment Loss or Pollution?</b></td>
			<td bgcolor="#999999"><select name="sediment">
					<option value="1"<% If rsReport("sediment")="1" Then %> selected<% End If %>>Yes</option>
					<option value="0"<% If rsReport("sediment")="0" Then %> selected<% End If %>>No</option>
				</select></td></tr>
<% END IF %>
		<TR><TD align="right" bgcolor="#eeeeee"><b>Narrative</b></td>
			<td bgcolor="#999999">
				<INPUT type="button" value="Edit Narrative" onClick="editNarrative('<%= inspecID%>');"></TD></TR>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><!-- Type of Report? Weekly, Storm, Complaint, ? --><tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
<%	'admin can change inspector name
If Session("validAdmin") Then
	insSQLSELECT = "SELECT DISTINCT u.userID, firstName, lastName" & _
		" FROM Users as u, ProjectsUsers as pu" & _
		" WHERE u.userID = pu.userID AND (pu.rights='inspector' OR pu.rights='admin')" &_
		" ORDER BY lastName"
	Set connUser = connSWPPP.execute(insSQLSELECT) %>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>Inspector:</strong></td>
			<td bgcolor="#999999"><select name="inspector">
				<% Do While Not connUser.EOF %><option value="<%= connUser("userID") %>" <% If rsReport("userID")=connUser("userID") Then %>selected<% End If %>><%= Trim(connUser("firstName")) & "&nbsp;" & Trim(connUser("lastName")) %></option> <%= rsReport("userID") %>*
<%					connUser.moveNext
				Loop				
	connUser.Close
	Set connUser = Nothing %>
			</select></td></tr>
<%	Else 
	SQLa="SELECT * FROM Users WHERE userID="& rsReport("userID") 
	Set connUser= connSWPPP.execute(SQLa) %>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>Inspector:</strong></td>
			<td bgcolor="#999999"><%= Trim(connUser("firstName"))%> <%=Trim(connUser("lastName"))%>
				<INPUT type="hidden" name="inspector" value="<%= rsReport("userID")%>"></td></tr>
<%	End If
	IF Session("validAdmin") OR Session("validInspector") THEN
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	Set objSteMapDir = folderSvrObj.GetFolder(baseDir & "images\sitemap\")
	Set siteMapImage = objSteMapDir.Files 

	SQLa="sp_oImagesByType "& inspecID &",12"
'response.write(SQLa)
	SET RSa=connSWPPP.execute(SQLa) 
	tempStrOfFileNames="" 
	t1="sitemap"
	t2="sitemapDN"
	t3="sitemapUP" %>
<!--		<tr><td align="right" bgcolor="#eeeeee"><strong>Site Map File:</strong></td>
			<td bgcolor="#999999" nowrap>
				<SPAN id="sitemapSPAN">
				<select name="sitemapDN" size="1" class="long">
<% 	DO WHILE NOT(RSa.EOF) %><OPTION value="<%= Trim(RSa("oImageFileName"))%>"><%= Trim(RSa("oImageFileName"))%></OPTION>
<%		tempStrOfFileNames=tempStrOfFileNames & TRIM(RSa("oImageFileName"))&":"
		RSa.MoveNext
	LOOP %>		</SELECT>
				<input type="hidden" name="sitemap" value="<%= tempStrOfFileNames%>">
					<BUTTON onClick="delOption(<%= t1%>, <%= t2%>, <%= t3%>);">--&gt;</BUTTON>
					<BUTTON onClick="addOption(<%= t1%>, <%= t2%>, <%= t3%>);">&lt;--</BUTTON>
				<select name="sitemapUP" class="long">
<%	For Each Item In siteMapImage
		shortName = Item.Name 
		IF InStr(tempStrOfFileNames, shortName)=0 THEN %><option value="<% = Trim(shortName) %>"><% = Trim(shortName) %></option>
<%		END IF
	Next
	Set objSteMapDir = Nothing
	Set siteMapImage = Nothing %>
				</select></SPAN> &nbsp;&nbsp; <input type="button" value="Upload Site Map File" 
					onClick="location='upSiteMapEditRprt.asp?inspecID=<% = inspecID %>'; return false";>
				</td></tr>-->
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><!-- Type of Report? Weekly, Storm, Complaint, ? --><tr> 
</Table>

<!-- ------------- Optional Links ----------------------------------------------------- -->

<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">
		<tr><td colspan="6" align="center"><font size="+1">Optional Project Links</font></td></tr>
		<tr><td colspan="5"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999" colspan=5><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr valign="bottom">
		<td align="right" bgcolor="#eeeeee" style="border-bottom: 1px solid black;"><font size="-2"><strong>Type</strong></font></td>
		<td bgcolor="#999999" style="padding: 0cm; margin: 0cm; border-bottom: 1px solid black;">&nbsp;</td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black; padding: 0cm; margin: 0cm;"><font size="-2"><strong>Current Links</strong></font></td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black;" align="center"><font size="-2"><strong><nobr>rem | add</nobr></strong></font></td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black;"><font size="-2"><strong>Available Links</strong></font></td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black;">&nbsp;</td></tr>
<%	
SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal >=0 ORDER BY oitSortByVal asc"
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
	Set TempImage = objTemp.Files 
	SQLa="sp_oImagesByType "& inspecID &",'"& RS1("oitID") &"'" 
	SET RSa=connSWPPP.Execute(SQLa)
	tempStrOfFileNames="" %>
		<tr valign="top">
			<td align="right" bgcolor="#eeeeee"><font size="-1"><strong><%= oitDesc%>:</strong></font></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right">
				<BUTTON onClick="swapOption(<%= t1%>, <%= t2%>, 'up');">&uarr;</BUTTON><br>
				<BUTTON onClick="swapOption(<%= t1%>, <%= t2%>, 'dn');">&darr;</BUTTON></td>	
			<td bgcolor="#999999" nowrap align="left" style="padding: 0cm; margin: 0cm;" align=left>
				<select name="<%=dirName%>DN" size="3" class="long" style="padding: 0cm; margin: 0cm;" align=left>
<% 	DO WHILE NOT(RSa.EOF) %><OPTION value="<%= Trim(RSa("oImageFileName"))%>"><%= Trim(RSa("oImageFileName"))%></OPTION>
<%		tempStrOfFileNames=tempStrOfFileNames & TRIM(RSa("oImageFileName"))&":"
		RSa.MoveNext
	LOOP %>		</SELECT></td>				
			<td bgcolor="#999999" align="center">
				<BUTTON onClick="delOption(<%= t1%>, <%= t2%>, <%= t3%>);">--&gt;</BUTTON>
				<BUTTON onClick="addOption(<%= t1%>, <%= t2%>, <%= t3%>);">&lt;--</BUTTON></td>
			<td bgcolor="#999999"><select name="<%=dirName%>UP"  class="long">
<%	For Each Item In TempImage
		shortName = Item.Name 
		IF InStr(tempStrOfFileNames, shortName)=0 THEN %><option value="<% = Trim(shortName) %>"><% = Trim(shortName) %></option>
<%		End If  '-- this weeds out those already selected fileNames
	Next 
	Set objTemp = Nothing
	Set TempImage = Nothing %>
				</select></td>
			<td bgcolor="#999999" nowrap align="right"><input type="button" style="width:150;"
				value="Upload File/Image" onClick="location='upImageRpt.asp?inspecID=<%= inspecID %>&oitID=<%= oitID %>'; return false";>
				<input type="hidden" name="<%= t1%>" value="<%= tempStrOfFileNames%>"></td></tr>
<% 	RS1.MoveNext
LOOP %>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"/></td>
			<td bgcolor="#999999" colspan=5><img src="../../images/dot.gif" width="5" height="5"/></td>
		</tr><tr><td colspan="5"><img src="../../images/dot.gif" width="5" height="5"/></td>
		</tr>
		<tr><td colspan="5">&nbsp;</td>
		</tr><tr><td colspan=6 align="center"><input name="submit" type="submit" value="Edit Report & Project Info"/></td></tr>
</table>

<% End If 'Session("validAdmin") %>

<!------------------------------------- Images ---------------------------------------->

<% IF NOT(Session("noImages")) THEN %>
	<h1>Images</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0"><%
smImgSQLSELECT = "SELECT imageID, smallImage, description" & _
	" FROM Images WHERE inspecID=" & inspecID	
Set rsSmImages = connSWPPP.execute(smImgSQLSELECT)

If rsSmImages.EOF Then
	Response.Write("<tr><td colspan='2' align='center'><i>There are " & _
		"no images at this time.</i></td></tr>")
Else %> 
	<tr><td colspan="3">Edit an image or description by selecting the name.<br><br></td></tr>
	<tr>
	<% Do While Not rsSmImages.EOF
	imageID = rsSmImages("imageID")
	smallImage = Trim(rsSmImages("smallImage"))
	desc = Trim(rsSmImages("description"))
	
	iDataRows = iDataRows + 1
	If iDataRows > 3 Then
		Response.Write("	</tr>" & VBCrLf & "	<tr>")
		iDataRows = 1
	End If %>
	<td height="125"><div align="center"><a href="editImage.asp?imageID=<%= imageID %>&inspecID=<%=inspecID%>"><%= desc %></a><br>
	<a href="editImage.asp?imageID=<%= imageID %>">
	<img src="../../images/sm/<%= smallImage %>" alt="<%= smallImage %>" width="100" height="75" 
		border="0"></a></div></td>
	<% rsSmImages.MoveNext
	Loop %>
	</tr>
	<% End If

	rsSmImages.Close 
	Set rsSmImages = Nothing %>
	<tr><td colspan="3" align="center"><br><input type="button" value="Add New Image" 
		onClick="location='addImage.asp?inspecID=<% = inspecID %>'; return false";></td></tr></table>
<% END IF	'--- noImages Check %>

<!------------------------------------- Coordinates --------------------------- --->
<h1>Location</h1>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
    <tr><td width="10%" nowrap>list order&nbsp;&nbsp;&nbsp;<img width="16" height="16" src="../../images/trash.ico"/></td><td></td><td></td></tr>
	<tr><td colspan="4"><hr align="center" width="100%" size="1"></td></tr>	<%
coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby" &_
	" FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
Set rsCoord = connSWPPP.execute(coordSQLSELECT)
'If rsCoord.EOF Then
'	Response.Write("<tr><td colspan='2' align='center'><i>There is no data at this time.</i></td></tr>")		
'Else
    n = 1
	Do While Not rsCoord.EOF	
	    coID = rsCoord("coID")
		correctiveMods = Trim(rsCoord("correctiveMods"))
		orderby = rsCoord("orderby")
		coordinates = Trim(rsCoord("coordinates"))
		existingBMP = Trim(rsCoord("existingBMP")) %>
	<tr>
		<td width="10%" rowspan="3" align="center">
			<input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
			<input type="text" name="coord:orderby:<%= n %>" size="4" value="<% = orderby %>" />&nbsp;
			<input type="checkbox" name="coord:del:<%= n %>" />
		</td>
		<td width="10%" align="right"><b>Location:</b></td>
		<td width="80%">
			<input name="coord:coord:<%= n %>" type="text" value="<%= coordinates %>" size="175">
		</td>
	</tr>
<%	IF existingBMP <> "-1" THEN %>
	<tr>
		<td align="right"><b>Existing BMP:</b></td>
		<td><font face="Times" size="2.5pt"><%= existingBMP %></font></td>
	</tr>
<% 	END IF %>
	<tr>
		<td align="right" valign="top" nowrap><b>Corrective Mods:</b><!--<br />(required)--></td><td>
	    <textarea name="coord:mods:<%= n %>" cols="150" rows="10"><%= correctiveMods %></textarea></td>
	</tr>
	<tr>
		<td colspan="3"><hr align="center" width="100%" size="1"></td>
	</tr>
<%	 	rsCoord.MoveNext
        n = n + 1
	Loop 	
'End If ' END No Results Found
rsCoord.Close
Set rsCoord = Nothing
connSWPPP.Close 
Set connSWPPP = Nothing 
%>	<tr>
		<td width="10%" rowspan="3" align="center">
	        <input type="hidden" name="coord:coID:0" value="0" />
	        <input type="text" name="coord:orderby:0" size="4" value=""/>&nbsp;
	        <input type="checkbox" name="coord:del:0" disabled="disabled" />
		</td>
		<td width="10%" align="right"><b>Location:</b></td>
		<td width="80%"><input name="coord:coord:0" type="text" value="" size="175"/></td>
	</tr>
	<tr>
		<td align="right" valign="top" nowrap><b>Corrective Mods:</b><!--<br />(required)--></td><td>
			<textarea name="coord:mods:0" cols="150" rows="10" ></textarea></td>
	</tr>
	<tr>
		<td colspan="3"><hr align="center" width="100%" size="1"></td>
	</tr>	
	<tr>
		<td colspan="3" align="center"><br>
	<input name="submit" type="submit" value="Modify Coordinates"/>
		<br><br></td></tr>
<br><br>
</form>
</body>
</html>