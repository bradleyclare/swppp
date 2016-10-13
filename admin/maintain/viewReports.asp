<%@  language="VBScript" %>
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
%><!-- #include virtual="admin/connSWPPP.asp" --><%
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
    <!-- #include virtual="admin/adminHeader2.inc" -->
	<div class="six columns alpha">
		<h3>View Inspection Reports</h3>
	</div>
    <div class="six columns omega">
		for the month of:
		<select id="startMonth" name="startMonth" onchange="navigateMe();">
        <% m=DateDiff("m",startDate,Date)+12
'	IF Month(startDate)=Month(Date) THEN m=-13 ELSE m=-8 END IF
		FOR n= -m to (m-12) step 1 
			optDate=DateAdd("m",n,startDate)
			optMonth=MonthName(Month(optDate))
			optYear=Year(optDate) %>
            <option value="<%= optDate%>" <% IF Month(optDate)=startMonth THEN%> selected<% END IF%>><%= optMonth%>, <%= optYear%> <%	Next %>
        </select>
    </div>
	<table width="100%" border="0">
        <tr>
            <th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=inspecDate DESC&startDate=<%=startDate%>"><b>Date</b></a></th>
            <th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=p.projectName, p.projectPhase, inspecDate DESC&startDate=<%=startDate%>"><b>Name</b></a></th>
            <th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=reportType&startDate=<%=startDate%>"><b>Type</b></a></th>
            <th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=released, p.projectName, p.projectPhase, inspecDate DESC&startDate=<%=startDate%>"><b>released</b></a></th>
            <th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=lastName&startDate=<%=startDate%>"><b>Inspector</b></a></th>
        </tr>
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
            <td><a href="editReport.asp?inspecID=<% = rsReports("inspecID") %>"><%= projectName %>&nbsp;<%= projectPhase%></a></td>
            <td><% = reportType %></td>
            <td>
                <img src="../../images/<% IF released THEN %>checkbox_1.gif<% ELSE %>checkbox_0.gif<% END IF %>"></td>
            <td><% = userFullName %></td>
        </tr>
        <%
			' Alternate Row Colors
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			
			rsReports.MoveNext
		Loop
		
	End If ' END No Results Found
	
rsReports.Close
Set rsReports = Nothing

connSWPPP.Close
Set connSWPPP = Nothing %>
    </table>
	</div>
    <script type="text/javascript">
        function navigateMe() {
            var select_obj = document.getElementById("startMonth");
            var id = select_obj.selectedIndex;
            var month = select_obj.options[id].text;
            var link = "viewReports.asp?startDate=" + month;
            window.open(link, "_self");
        }
    </script>
</body>
</html>
