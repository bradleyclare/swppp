<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

startDate=Request("startDate")
'IF NOT(IsDate(startDate)) THEN Response.Write("not a date<br>") END IF
IF (TRIM(startDate)="" OR NOT(IsDate(startDate)) )Then 
	startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
Else
	startDate=CDATE(Month(CDATE(Request("startDate")))&"/1/"&Year(CDATE(Request("startDate"))))
End IF
endDate=DateAdd("m",1,startDate)
endDate=DateAdd("d",-1,endDate)
startMonth=Month(startDate)
startYear=Year(startDate)
recordOrd = Request("orderBy")
If recordOrd = "" Then recordOrd = "p.projectName, p.projectPhase, inspecDate DESC" End If
%><!-- #include file="../connSWPPP.asp" --><%
reportSQLSELECT = "SELECT inspecID, inspecDate, reportType" & _
	", firstName, lastName, i.projectID, i.projectName, i.projectPhase, released" & _
	" FROM Inspections as i, Users as u, Projects as p" & _
	" WHERE i.userID = u.userID AND i.projectID = p.projectID" &_
	" AND i.inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" 
	IF session("validInspector") AND NOT(Session("validAdmin")) THEN reportSQLSELECT = reportSQLSELECT & " AND i.userID='" & Session("userID") &"'"
	reportSQLSELECT = reportSQLSELECT & " ORDER BY " & recordOrd
'Response.Write(reportSQLSELECT & "<br>")
Set rsReports = connSWPPP.execute(reportSQLSELECT)
%>
<html>
<head>
<title>SWPPP INSPECTIONS : View Inspection Reports</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<table width="90%" border="0">
	<tr> 
		<TD align=left colspan=2><h1>reports</h1></TD>
		<TD align=right colspan=3>for the month of:&nbsp;<select id = "startMonth" name="startMonth" 
			onchange="navigateMe();">
<%	
	m=DateDiff("m",startDate,Date)+12
'	IF Month(startDate)=Month(Date) THEN m=-13 ELSE m=-8 END IF
	FOR n= -m to (m-12) step 1 
		optDate=DateAdd("m",n,startDate)
		optMonth=MonthName(Month(optDate))
		optYear=Year(optDate) %>
			<OPTION value="<%= optDate%>"<% IF Month(optDate)=startMonth THEN%> selected<% END IF%>><%= optMonth%>, <%= optYear%>
<%	Next %>
			</SELECT>
		</TD>
	</tr>
	<tr><th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=inspecDate DESC&startDate=<%=startDate%>"><b>date</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=p.projectName, p.projectPhase, inspecDate DESC&startDate=<%=startDate%>"><b>name</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=reportType&startDate=<%=startDate%>"><b>type</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=released, p.projectName, p.projectPhase, inspecDate DESC&startDate=<%=startDate%>"><b>released</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=lastName&startDate=<%=startDate%>"><b>inspector</b></a></th></tr>
	<tr><td colspan="5"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
<%
	If rsReports.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>Sorry " & _
			"no current data entered at this time.</i></b></td></tr>")
	Else
		altColors = "#e5e6e8"
		
		Do While Not rsReports.EOF
			inspecDate = rsReports("inspecDate")
			projectName = Trim(rsReports("projectName"))
			projectPhase = Trim(rsReports("projectPhase"))
			reportType = Trim(rsReports("reportType"))
			released = Trim(rsReports("released"))
			userFullName = Trim(rsReports("firstName")) & "&nbsp;" & Trim(rsReports("lastName"))
%>
	<tr align="center" bgcolor="<% = altColors %>"> 
		<td><% = inspecDate %></td>
		<td nowrap><a href="editReport.asp?inspecID=<% = rsReports("inspecID") %>"><%= projectName %>&nbsp;<%= projectPhase%></a></td>
		<td nowrap><% = reportType %></td>
		<td nowrap><img src="../../images/<% IF released THEN %>checkbox_1.gif<% ELSE %>checkbox_0.gif<% END IF %>"></td>
		<td nowrap><% = userFullName %></td></tr>
<%
			' Alternate Row Colors
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			
			rsReports.MoveNext
		Loop
		
	End If ' END No Results Found
	
rsReports.Close
Set rsReports = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
	<tr><td colspan="5">&nbsp;</td></tr>
</table>
<script type="text/javascript">
function navigateMe(){
	var select_obj = document.getElementById("startMonth");
	var id = select_obj.selectedIndex;
	var month = select_obj.options[id].text;
	var link = "viewReports.asp?startDate=" + month;
	window.open(link,"_self");
}
</script>
</body>
</html>