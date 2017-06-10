<%@ Language="VBScript" %>
<%
If 	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") _	
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")	
End If
'--	Determine the month to report on -----------------------------------------------------------
startDate=Trim(Request("startDate"))
endDate=Trim(Request("endDate"))
IF (startDate="" OR NOT(IsDate(startDate)) )Then startDate=CDATE(Month(Date) &"/1/"& Year(Date)) End IF
IF (TRIM(endDate)="" OR NOT(IsDate(endDate)) ) THEN 
	endDate=DateAdd("d",-1,(DateAdd("m",1,startDate))) 
	monthRange=true
ELSE
	monthRange=false
END IF
IF CDATE(endDate)< CDATE(startDate) THEN endDate=DateAdd("d",-1,(DateAdd("m",1,startDate))) END IF
startMonth=Month(startDate)
startYear=Year(startDate)
recordOrd = Request("orderBy")
If recordOrd = "" Then recordOrd = " p.projectName ASC, p.projectPhase, inspecDate DESC" End If
%><!-- #include file="../admin/connSWPPP.asp" --><%
SQL0 = "SELECT inspecID, inspecDate, projectCounty" & _
	", i.projectID, p.projectName, p.projectPhase, p.collectionName" & _
	" FROM Inspections as i, Projects as p" & _
	" WHERE i.projectID = p.projectID" &_
	" AND i.inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
	" AND i.projectID IN ( SELECT projectID FROM ProjectsUsers" 
IF NOT Session("validAdmin") THEN
	SQL0=SQL0 & " WHERE  ProjectsUsers.userID = " & Session("userID") 
END IF
	SQL0=SQL0 & ") ORDER BY " & recordOrd
Set RS0 = connSWPPP.execute(SQL0) %>
<html>
<head>
<title>SWPPP INSPECTIONS : View Inspection Reports</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
<STYLE>
    .visYes { visibility:visible; display: inline }
    .visNo { visibility:hidden; display: none}
</STYLE>

</head>
<body>
<!-- #include file="../header2.inc" -->
<table width="90%" border="0">
	<tr valign="middle"> 
		<TD align=left colspan=5><h1>View Inspection Reports</h1></TD></tr>
<% If Session("validAdmin") then %>
	<tr valign=top><TD align=right colspan=3 valign="middle">
			<FORM action="qbXLS.asp" method="post">
				Create Spreadsheet for <%= MonthName(startMonth)%> with starting invoice number 
				<input type="text" name="iNum" value="" maxlength="6" size="4">
				<input type="hidden" name="xDate" value="<%= startDate%>">
				<input type="hidden" name="yDate" value="<%= endDate%>"><br>
				and billing cycle <SELECT name="bCycle">
						<OPTION value="1" selected>1</option>
						<OPTION value="2">2</option>
						<OPTION value="3">3</option>
						<OPTION value="4">4</option></SELECT>
				<input type="submit" value="GO">&nbsp;
				<button style="border-width:thin; font-size:xx-small; height:17px; width:17px; background-color:#e5e6e8;" 
					onClick="alert('IMPORT HELP\n\n1. Enter the next invoice number from Quickbooks.When you click the \'go\' button\nthe system will create a spreadsheet that you can import into Quickbooks.\n\n2. \'Save As\' the generated spreadsheet as a tab delimited file. Use any file-name\nyou need, but enclose the file-name in quotation marks and end the file-name\nwith \'.iff\' before saving. This will allow the file to be imported into Quickbooks.\n   (for example: \'\'testfile.iff\'\')\n\n3. In Quickbooks, import the file by accessing \'File>>Utilities>>Import IIF file\'.\nThis will import the invoices into the system.');">?</button></FORM></TD>
		<TD align=right colspan=2>
<% ELSE %><TD align=right colspan=5><% END IF %>View Reports for 
			<SELECT name="repType" onChange="spans(this.value)">
				<OPTION value="monthX" <% if monthRange then %>selected<% end if %>>the month of</option>
				<option value="dateX"<% if not(monthRange) then%> selected<% end if%>>date range</option>
			</SELECT><br>
			<span id="span1" class="<% if monthRange then %>visYes<% else %>visNo<% end if %>">
			<SELECT name="startMonth" id="startMonth" onChange="navigateMe();"><%	
	m=DateDiff("m",startDate,Date)+6
'	IF Month(startDate)=Month(Date) THEN m=-13 ELSE m=-8 END IF
	FOR n= -m to (m-6) step 1 
		optDate=DateAdd("m",n,startDate)
		optMonth=MonthName(Month(optDate))
		optYear=Year(optDate) %>
				<OPTION value="<%= optDate%>"<% IF Month(optDate)=startMonth THEN%> selected<% END IF%>><%= optMonth%>, <%= optYear%>
<%	Next %>	</SELECT></span>
			<span class="<% if monthRange then %>visNo<% else %>visYes<% end if %>" id="span2">
			from: <input type="text" id="startDate" name="startDate" value="<%= startDate %>" onBlur="validDate(this)"><br>
			to: <input type=text id="endDate" name="endDate" value="<%= endDate %>" onBlur="validDate(this)"><br>
			<button style="border-width:thin; font-size:xx-small; height:20px; width:26px; background-color:#e5e6e8;"
				onClick="navigate();">GO</button></span></TD></tr>
	<tr><!--		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=projectCounty&startDate=<%=startDate%>"><b>County</b></a></th>-->
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=p.projectName, p.projectPhase, inspecDate DESC&startDate=<%=startDate%>"><b>Project</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=collectionName, p.projectName, p.projectPhase DESC&startDate=<%=startDate%>"><b>Group</b></a></th>
        <th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=inspecDate DESC&startDate=<%=startDate%>"><b>Date</b></a></th></tr>
	<tr><td colspan="3"><img src="../../images/dot.gif" width="5" height="5"></td></tr><%	
	If RS0.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>Sorry " & _
			"no current data available for the Month requested.</i></b></td></tr>")
	Else
		altColors = "#e5e6e8"		
		Do While Not RS0.EOF
			inspecID = RS0("inspecID")
			inspecDate = RS0("inspecDate")
			projectName = Trim(RS0("projectName"))
			projectPhase = Trim(RS0("projectPhase"))
			projectCounty = Trim(RS0("projectCounty"))
            groupName = Trim(RS0("collectionName"))%>
	<tr align="center" bgcolor="<% = altColors %>" onMouseOver="this.bgColor='#006699';" onMouseOut="this.bgColor='<%=altColors%>';" onClick="window.location='report.asp?inspecID=<%= inspecID%>';"> 
<!--		<td nowrap><% = projectCounty %></td>-->
		<td nowrap><% = projectName %>&nbsp<%= projectPhase%></td>
		<td><% = groupName %></td>
        <td><% = inspecDate %></td>
	</tr><%
			' Alternate Row Colors
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If			
			RS0.MoveNext
		Loop		
	End If ' END No Results Found	
RS0.Close
Set RS0 = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
	<tr><td colspan="3">&nbsp;</td></tr>
</table>
</body>
<script type="text/javascript">
function navigateMe(){
	var select_obj = document.getElementById("startMonth");
	var id = select_obj.selectedIndex;
	var month = select_obj.options[id].text;
	var link = "monthlyReportsSum.asp?startDate=" + month;
	window.open(link,"_self");
}

function navigate(){
	var start = document.getElementById("startDate").value;
	var end = document.getElementById("endDate").value;
	var link = "monthlyReportsSum.asp?startDate=" + start + "&endDate=" + end;
	window.open(link,"_self");
}

<!-- 
function spans(rType){
	if (rType=="dateX") {
		document.all.span2.className="visYes";
		document.all.span1.className="visNo";
	}	
	else {
		document.all.span2.className="visNo";
		document.all.span1.className="visYes";
	}

}
function validDate(objX){
	var test = new Date(objX.value);
//	alert(test.getMonth());
	if ( Boolean(y2k(test.getYear()))&& test.getMonth()>=0 && test.getMonth()<=11 && Boolean(test.getDate())){
		return true;}
	else {
		objX.value="";
		objX.focus();			
      	return false;}
}

function y2k(number) { return (number < 100) ? number + 2000 : number; }
-->
</script>
</html>