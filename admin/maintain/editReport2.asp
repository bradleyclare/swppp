<%Response.Buffer = False%>
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
%><!-- #include file="../connSWPPP.asp" --><%
If Request.Form.Count > 0 Then	
	'Response.Write("Form Submitted")
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	
	if Request.Form("submit_optional_btn") = "Modify Optional Links" then
		Response.Redirect("editReportOptionalLinks.asp?inspecID=" + inspecID)
	End If
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
		includeItems = 0
		if Request("includeItems") = "on" then includeItems = 1 End If
		compliance = 0
		if Request("compliance") = "on" then compliance = 1 End If
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
		", includeItems = " & includeItems & _
		", compliance = " & compliance & _
		" WHERE inspecID = " & inspecID
'response.Write(inspectSQLUPDATE)
	connSWPPP.Execute(inspectSQLUPDATE)
    
		totalItems = 0
		completedItems = 0
		for n = 1 to 999 step 1
'Response.Write("coord:coID:" & CStr(n)&":"& Request("coord:coID:" & CStr(n)) &"<br/>")
		    if Trim(Request("coord:coID:" & CStr(n))) = "" then
		        exit for
		    end if
'-- dbo.spAEDCoordinate @_iCOID int, @_DelFlag smallint, @_inspecID int, @_iCoordinates char(50), @_icorrectiveMods char(255), @_iOrderBy int
            DelCoord = 0
            if Request("coord:del:"& CStr(n)) = "on" then 
				DelCoord = 1 
			Else
				totalItems = totalItems + 1
			End If
			Complete = 0
			if Request("coord:status:"& CStr(n)) = "on" then 
				Complete = 1 
				completedItems = completedItems + 1
			End If
			Repeat = 0
			if Request("coord:repeat:"& CStr(n)) = "on" then Repeat = 1 End If
			useAddress = 0
            if Request("coord:useAddress:"& Cstr(n)) = "on" then useAddress = 1 End If
			address = TRIM(strQuoteReplace(Request("coord:addressName:"& Cstr(n))))
			locationName = TRIM(strQuoteReplace(Request("coord:locationName:"& Cstr(n))))
			AssignDate = date()
			if Repeat = 1 Then
				AssignDate = Request("coord:assignDate:"& CStr(n))
			End If
			OrderBy = 0
            if IsNumeric(Request("coord:orderby:"& CStr(n))) then OrderBy = Request("coord:orderby:"& CStr(n)) End If
			'SQLc = SQLc &"/*<br/>*/Exec spAEDCoordinate "& Request("coord:coID:"& CStr(n)) &", "& DelCoord &", "& inspecID &", '"& Replace(Request("coord:coord:"& CStr(n)),"--","�") &"', '"& Replace(Request("coord:mods:"& CStr(n)),"--","�") &"', "& OrderBy &";"
			SQLc = SQLc &"/*<br/>*/Exec spAEDCoordinate "& _ 
			Request("coord:coID:"& CStr(n)) &", "& _ 
			DelCoord &", "& _ 
			inspecID &", '"& _ 
			Replace(Request("coord:coord:"& CStr(n)),"--","�") &"', '"& _ 
			Replace(Request("coord:mods:"& CStr(n)),"--","�") &"', "& _ 
			OrderBy &", '"& _ 
			AssignDate &"', '"& _ 
			Request("coord:completeDate:"& CStr(n)) &"', "& _ 
			Complete &", " & _ 
			Repeat &", " & _ 
			useAddress &", '" & _ 
			address &"', '" & _
			locationName &"';"
		next	
'Response.Write(SQLc)
        if Len(SQLc) > 0 then connSWPPP.execute(SQLc) end if
'Response.End	

	'update items counts
	inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
		" totalItems = " & totalItems & _
		", completedItems = " & completedItems & _
		" WHERE inspecID = " & inspecID
'response.Write(inspectSQLUPDATE2)
	connSWPPP.Execute(inspectSQLUPDATE2)

    if Request.Form("submit_view_reports_btn") = "View Reports" then
		connSWPPP.Close
	    Set connSWPPP = Nothing
    	Response.Redirect("viewReports.asp")
	End If
	If request("submit") = "Edit Report & Project Info" Then	
	    connSWPPP.Close
	    Set connSWPPP = Nothing
    	Response.Redirect("viewReports.asp")
    End If
End If
	inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, projectAddr, projectCity, projectState" & _
		", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, i.projectID, compName" & _
		", compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax" & _
		", contactEmail, reportType, inches, bmpsInPlace, sediment, userID, includeItems, compliance, totalItems, completedItems" & _
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
	<link href="../../css/jquery-ui.min.css" rel="stylesheet" type="text/css"/>
	<link href="../../css/jquery-ui.structure.min.css" rel="stylesheet" type="text/css"/>
	<link href="../../css/jquery-ui.theme.min.css" rel="stylesheet" type="text/css"/>
	<script src="../../js/jquery.js" type="text/javascript"></script>
	<script src="../../js/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" >
  $( function() {
    $( ".datepicker" ).datepicker();
  } );
</script>
<script type="text/javascript" >
$(document).ready( function () {
	$('#dialog-confirm').dialog({
      autoOpen: false,
	  resizable: false,
      height: "auto",
      width: 500,
      modal: true,
      buttons: {
        "Complete All Items": function() {
		  //check all complete hidden elements (coord:statusX)
		  var i;
		  for (i = 1; i < 99; ++i) {
				var e = document.getElementsByName("coord:status:" + i);
				if (e.length){
					$("[name='coord:status:" + i + "']")[0].checked = true;
				} else {
					break;
				}
		  }
		  $('#compliance-checkbox')[0].checked = true;
		  $(this).dialog("close");
		  document.getElementById("theForm").submit();
        },
		"Delete All Items": function() {
		  //check all delete checkboxes (coord:deleteX)
		  var i;
		  for (i = 1; i < 99; ++i) {
				var e = document.getElementsByName("coord:del:" + i);
				if (e.length){
					$("[name='coord:del:" + i + "']")[0].checked = true;
				} else {
					break;
				}
		  }
		  $('#compliance-checkbox')[0].checked = true;
		  $(this).dialog("close");
		  document.getElementById("theForm").submit();
        },
        Cancel: function() {
		  $(this).dialog("close");
        }
      }
    });
	
	$('#compliance-checkbox').click(
		function () {
			if ($('#compliance-checkbox').is(":checked")){
				$("#dialog-confirm").dialog('open');
                return false;
			} else {
				$('#compliance-checkbox')[0].checked = false;
			}
		}
	);
	
	$('#includeItems-checkbox').click(
		function () {
			document.getElementById("theForm").submit();
		}
	);
});
</script>
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

function useAddressLookup(obj){
	var parts = obj.name.split(":");
	var selectname = "coord:addOptions:" + parts[2];
	var s = document.getElementsByName(selectname);
	var selectname2 = "coord:address:" + parts[2];
	var s2 = document.getElementsByName(selectname2);
	var selectname3 = "coord:coord:" + parts[2];
	var s3 = document.getElementsByName(selectname3);
	if (obj.checked) //enable select object
	{
		s[0].className = "";
		s2[0].className = "";
		s3[0].className = "hide";
	}
	else //disable select object
	{
		s[0].className = "hide";
		s2[0].className = "hide";
		s3[0].className = "";
	}
}

function setSelectValue(obj){
	//selected value of addOptions dropdown
	var val         = obj.selectedIndex;
	var parts       = obj.name.split(":");
	
	//find address dropdown list to set the same value
	var selectname     = "coord:address:" + parts[2];
	var s              = document.getElementsByName(selectname);
	s[0].selectedIndex = val;
	
	//set the hidden object to keep address name
	var hiddenname2 = "coord:addressName:" + parts[2];
	var s2          = document.getElementsByName(hiddenname2);
	s2[0].value     = s[0].value.trim();
	
	//set the hidden object for locationName
	var hiddenname3 = "coord:locationName:" + parts[2];
	var s3          = document.getElementsByName(hiddenname3);
	s3[0].value     = obj.value.trim();
}

</script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Edit Inspection Report</h1>
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
	<input type="hidden" name="inspecID" value="<%=inspecID %>"/>
	<input type="hidden" name="projectID" value="<%=rsReport("projectID") %>"/>
	
<div id="dialog-confirm" title="My Dialog Title">
	<p>Site in Compliance? What do you want to do with the open items?</p>
</div>
<table width="90%">
<tr><th width="30%" align="center">Report Date</th><th width="30%" align="center">Project Name</th><th width="30%" align="center">Customer Name</th></tr>
<tr><td align="center"><% = Trim(rsReport("inspecDate")) %></td><td align="center"><% = Trim(rsReport("projectName")) %></td><td align="center"><% = Trim(rsReport("compName")) %></td></tr>
<tr>
<td align="center"><a href="deleteReport.asp?inspecID=<%=inspecID %>"><button type="button">Delete Report</button></a></td>
<td align="center"><a href="releasereports_test.asp?inspecID=<%=inspecID%>&projID=<%=rsReport("projectID")%>"><button type="button">View Email Report</button></a></td>
<td align="center"><a href="manage_addresses.asp?inspecID=<%=inspecID%>"><button type="button">Upload Address Information</button></a></td>
</tr></table>
<br/>
<hr>
<table width="90%" border="0" cellpadding="2" cellspacing="0">
	<tr><td align="right" bgcolor="#eeeeee"><b>Type of Report:</b></td>
			<td bgcolor="#999999"><select name="reportType">
<% 	SQL2="SELECT * FROM ReportTypes where priority > 0 ORDER BY priority DESC, reportTypeID ASC"
	SET RS2=connSWPPP.execute(SQL2)
	DO WHILE NOT RS2.EOF %><option value="<%= Trim(RS2("reportType"))%>"<% 
	If Trim(rsReport("reportType")) = TRIM(RS2("reportType")) Then %> selected<% End If %>><%= Trim(RS2("reportType"))%></option>
<% 	RS2.MoveNext
	LOOP %>	</select></td>
		</tr>
	<TR><TD align="right" bgcolor="#eeeeee"><b>Narrative</b></td>
	<td bgcolor="#999999">
	<INPUT type="button" value="Edit Narrative" onClick="editNarrative('<%= inspecID%>');"></TD></TR>
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
<%	End If %>
</table>

<!------------------------------------- Coordinates --------------------------- --->
<hr/>
<h2>Action Items</h2>
<% totalItems = rsReport("totalItems")
completedItems = rsReport("completedItems")
if totalItems <> "" and totalItems <> 0 Then
	score = FormatNumber((completedItems/totalItems)*100,0) & "%" 
Else
	score = "N/A"
End If%>
<table width="90%">
<tr width="20%"><td>Total Items: <%=totalItems%></td><td width="20%">Completed Items: <%=completedItems%></td><td width="20%">Report Score:<%=score%></td><td width="20%">Site is in Compliance
<% If rsReport("compliance") = True Then %>
	<input id="compliance-checkbox" type="checkbox" name="compliance" checked/>
<% Else %>
	<input id="compliance-checkbox" type="checkbox" name="compliance" />
<% End If %>
</td><td width="20%" align="left">Apply Scoring to Report
<% If rsReport("includeItems") = True Then %>
	<input id='includeItems-checkbox' type="checkbox" name="includeItems" checked/>
<% Else %>
	<input id='includeItems-checkbox' type="checkbox" name="includeItems" />
<% End If %>
</td></tr></table><br/>
<center>Click "Repeat" on all items that you want the assign date to stay the same. All other items will be updated to the current date on SUBMIT.</center><br/>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
<% coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName" &_
	" FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)
addressSQLSELECT = "SELECT addressID, locationName, address FROM Addresses WHERE projectID=" & rsReport("projectID") & " ORDER BY locationName"
'Response.Write(addressSQLSELECT)
Set rsAddress = connSWPPP.execute(addressSQLSELECT)
'If rsCoord.EOF Then
'	Response.Write("<tr><td colspan='2' align='center'><i>There is no data at this time.</i></td></tr>")		
'Else
    n = 1
	currentDate = date()
	Do While Not rsCoord.EOF
	    coID = rsCoord("coID")
		correctiveMods = Trim(rsCoord("correctiveMods"))
		orderby = rsCoord("orderby")
		coordinates = Trim(rsCoord("coordinates"))
		existingBMP = Trim(rsCoord("existingBMP")) 
		assignDate = rsCoord("assignDate") 
		completeDate = rsCoord("completeDate")
		status = rsCoord("status")
		repeat = rsCoord("repeat")
		useAddress = rsCoord("useAddress")
		address = TRIM(rsCoord("address"))
		locationName = TRIM(rsCoord("locationName"))
		'Response.Write("ID: " & coID & ", Coord: " & coordinates & ", LocName: " & locationName & ", address: " & address & ", Mods: " & correctiveMods & "<br/>") 
		Do While Not rsAddress.EOF
			name1 = TRIM(rsAddress("locationName")) 
			addname1 = TRIM(rsAddress("address"))
			Exit Do
		Loop
		%>
	<input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
	<input type="hidden" name="coord:status:<%= n %>" value="<%= status %>" />
	<input type="hidden" name="coord:completeDate:<%= n %>" value="" />
	<tr><td>ID#</td>
	<td><%= coID %></td>
	<td>Address<input type="checkbox" name="coord:useAddress:<%= n %>" onclick="useAddressLookup(this)" 
	<% if (useAddress) = True Then %>
		 checked
	<% End If %>
	/></td>
	<td>
	<select name="coord:addOptions:<%= n %>" onchange="setSelectValue(this)" 
	<% if (useAddress) = False Then %>
		class="hide"
	<% End If %>
	>
	<% selname = name1
	if not rsAddress.EOF Then
	Do While Not rsAddress.EOF 
		name = TRIM(rsAddress("locationName")) 
		addname = TRIM(rsAddress("address"))
		if StrComp(address,addname) = 0 Then 
			selname = name %>
			<option value="<%=name%>" selected><%=name%></option>
		<% Else %>
			<option value="<%=name%>"><%=name%></option>
		<% End If 
		rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst 
	End If %>
	</select></td>
	<input type="hidden" name="coord:locationName:<%= n %>" value="<%=selname%>" />
	<td><select name="coord:address:<%= n %>" disabled  
	<% if (useAddress) = False Then %>
		class="hide"
	<% End If %>
	>
	<% selname = addname1 
	if not rsAddress.EOF Then
	Do While Not rsAddress.EOF 
		addname = TRIM(rsAddress("address"))
		if StrComp(address,addname) = 0 Then 
			selname = addname %>
			<option value="<%=addname%>" selected><%=addname%></option>
		<% Else %>
			<option value="<%=addname%>"><%=addname%></option>
		<% End If
		rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst
	End If %>
	</select></td></tr>
	<input type="hidden" name="coord:addressName:<%= n %>" value="<%=selname%>" />
	<tr><td>Order</td>
	<td><input type="text" name="coord:orderby:<%= n %>" size="10" value="<% = orderby %>" /></td>
	<td>Location Info:</td>
	<td colspan="2"><input name="coord:coord:<%= n %>" type="text" value="<%= coordinates %>" size="100%"  
	<% if (useAddress) = True Then %>
		class="hide"
	<% End If %>
	></td></tr>
	<tr><td>Delete<input type="checkbox" name="coord:del:<%= n %>" /></td>
	<td>Repeat
	<% If repeat = True Then %>
		<input type="checkbox" name="coord:repeat:<%= n %>" checked/>
	<% Else %>
		<input type="checkbox" name="coord:repeat:<%= n %>" />
	<% End If %>
	</td><td>Modifications:</td>
	<td rowspan="3" colspan="2"><textarea name="coord:mods:<%= n %>" cols="100%" rows="5"><%= correctiveMods %></textarea></td></tr>
	<tr><td>AssignDate</td>
	<td><input class=datepicker type="text" name="coord:assignDate:<%= n %>" size="10" value="<%= assignDate %>" /></td></tr>
	<tr><td></td><td></td></tr>
<%	IF existingBMP <> "-1" THEN %>
	<tr>
		<td align="right"><b>Existing BMP:</b></td>
		<td><font face="Times" size="2.5pt"><%= existingBMP %></font></td>
	</tr>
<% 	END IF %>
	<tr><td colspan="5"><hr align="center" width="100%" size="1"></td></tr>
<%	 	rsCoord.MoveNext
        n = n + 1
	Loop 	
'End If ' END No Results Found
rsCoord.Close
Set rsCoord = Nothing %>

<% for m = n to n+4 step 1 %>
	<input type="hidden" name="coord:coID:<%= m %>" value="0" />
	<input type="hidden" name="coord:del:<%= m %>" value="0" />
	<input type="hidden" name="coord:completeDate:<%= m %>" value="" />
	<input type="hidden" name="coord:status:<%= m %>" value="0" />
	<input type="hidden" name="coord:repeat:<%= m %>" value="0" />
	<tr><td>ID#</td>
	<td>0</td>
	<td>Address<input type="checkbox" name="coord:useAddress:<%= m %>" onclick="useAddressLookup(this)"/></td>
	<td><select name="coord:addOptions:<%= m %>" onchange="setSelectValue(this)" class="hide">
	<% selname = name1
	if not rsAddress.EOF Then
	Do While Not rsAddress.EOF 
		name = TRIM(rsAddress("locationName")) 
		If m = n Then
			selname = name
		End If %>
		<option value="<%=name%>"><%=name%></option>
	<% rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst
	End If %>
	</select></td>
	<input type="hidden" name="coord:locationName:<%= m %>" value="<%=selname%>" />
	<td><select name="coord:address:<%= m %>" disabled class="hide">
	<%	selname = addname1 
	if not rsAddress.EOF Then
	Do While Not rsAddress.EOF 
		name = TRIM(rsAddress("address")) 
		If m = n Then
			selname = name
		End If %>
		<option value="<%=name%>"><%=name%></option>
	<% rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst
	End If %>
	</select></td></tr>
	<input type="hidden" name="coord:addressName:<%= m %>" value="<%=selname%>" />
	<tr><td>Order</td>
	<td><input type="text" name="coord:orderby:<%= m %>" size="10" value="" /></td>
	<td>Location:</td>
	<td colspan="2"><input name="coord:coord:<%= m %>" type="text" value="" size="100%" ></td></tr>
	<tr><td></td>
	<td></td>
	<td>Mods:</td>
	<td rowspan="3" colspan="2"><textarea name="coord:mods:<%= m %>" cols="100%" rows="5"></textarea></td></tr>
	<tr><td>AssignDate</td>
	<td><input class=datepicker type="text" name="coord:assignDate:<%= m %>" size="10" value="<%= currentDate %>" /></td></tr>
	<tr><td></td><td></td></tr>
	<tr><td colspan="5"><hr align="center" width="100%" size="1"></td></tr>
<% next %>
	<tr><td colspan="5" align="center"><br>
	<input name="submit_coord_btn" type="submit" style="font-size: 20px;" value="Submit"/>
	<br><br></td></tr>
</table>
<% rsAddress.Close 
Set rsAddress = Nothing %>

<hr>
<h2>Project Information: #<%=rsReport("projectID")%></h2>
<table width="90%" border="0" cellpadding="2" cellspacing="0">
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
</table>
		
<h2>Company Information</h2>
<table width="90%" border="0" cellpadding="2" cellspacing="0">
		<tr><td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
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
		<!-- Type of Report? Weekly, Storm, Complaint, ? -->
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

	<% IF Session("validAdmin") OR Session("validInspector") THEN
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
</Table>

<!-- ------------- Optional Links ----------------------------------------------------- -->

<hr/>
<center><input name="submit_optional_btn" type="submit" style="font-size: 20px;" value="Modify Optional Links"/></center>
<hr/>
<center><input name="submit_view_reports_btn" type="submit" style="font-size: 20px;" value="View Reports"/></center>

<% End If 'Session("validAdmin") %>

<!------------------------------------- Images ---------------------------------------->

<% IF NOT(Session("noImages")) THEN %>
	<hr/>
	<h2>Images</h2>
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
	<tr><td colspan="3" align="center"><br><input type="button" style="font-size: 20px;" value="Add New Image" 
		onClick="location='addImage.asp?inspecID=<% = inspecID %>'; return false";></td></tr></table>
<% END IF	'--- noImages Check %>
</form>
<hr>
<% connSWPPP.Close 
Set connSWPPP = Nothing %>	
</body>
</html>