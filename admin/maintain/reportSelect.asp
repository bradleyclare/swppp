<%@ Language="VBScript" %>
<% Response.Buffer = False
If Not Session("validAdmin") and not session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

recordOrd = Request("orderBy")
If recordOrd = "" Then recordOrd = "inspecDate" End If
%>
<!-- #include file="../connSWPPP.asp" -->
<html>
<head>
<title>SWPPP INSPECTIONS : Select New or Default Inspection Report</title>
<link rel="stylesheet" type="text/css" href="../../global.css"/>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->

<%
currentDate = date()
endDate = currentDate
startDate=DateAdd("d",-90,currentDate)
SQL0=" SELECT i.inspecID, i.inspecDate, i.projectName, i.projectPhase, p.projectID "&_
        " FROM Inspections i inner join Projects p on i.projectid = p.projectid WHERE i.inspecDate Between '"& startDate &"' AND '"& endDate &"'" '&_
		'"   inner join (Select projectID, MAX(inspecDate) inspecDate From Inspections Group By projectID) as i2 " &_
        '"       on i.projectID = i2.projectID and i.inspecDate = i2.inspecDate and DateDiff(mm, i.inspecDate, GetDate()) < 3"

IF session("validInspector") AND NOT(Session("validAdmin")) THEN SQL0 = SQL0 & " AND i.userID='" & Session("userID") &"'"
SQL0 = SQL0 &" ORDER BY i.projectName asc, i.projectPhase"	
Response.Write("SQL0=" & SQL0)
SET rsReports2=connSWPPP.Execute(SQL0)
%>

<table width="100%" border="0">
	<form method="post" name="form_new_report" action="newReport.asp">
		<tr><td colspan="3" align="center"><input type="submit" value="add new project"/></td>
	</form>
	</tr><tr><td colspan="3"><img alt="" src="../../images/dot.gif" width="5" height="5" /></td></tr>
	<form method="post" name="form1" action="addReport.asp">
		<tr><td colspan="3">&nbsp;</td>
		</tr><tr><th width="15%"><a href="<%= Request.ServerVariables("script_name") %>?orderBy=inspecDate"><b>date</b></a></th>
			<th width="1080"><a href="<%= Request.ServerVariables("script_name") %>?orderBy=projectName"><b>report</b></a></th>
			<th width="16%"><input type=submit value="add report" /></th>
		</tr><tr><td colspan="3"><img alt="" src="../../images/dot.gif" width="5" height="5"/></td>
		</tr>
		<% If rsReports2.EOF Then
			Response.Write("<tr><td colspan='3' align='center'><b><i>Sorry " & _
				"no current data entered at this time.</i></b></td></tr>")
		Else
			altColors = "#e5e6e8"
			prevProjID = 0
			Do While Not rsReports2.EOF
				inspecID = rsReports2("inspecID")
				inspecDate = Trim(rsReports2("inspecDate"))
				projectName = Trim(rsReports2("projectName"))
				projectPhase = Trim(rsReports2("projectPhase"))
				projectID = rsReports2("projectID")
				If projectID <> prevProjID Then 
					prevProjID = projectID %>
					<tr nowrap><td align="center" bgcolor="<% = altColors %>"><% = inspecDate %></td>
					<td align="center" bgcolor="<% = altColors %>"><% = projectName %>&nbsp;<% = projectPhase %></td>
					<td align="center" bgcolor="<% = altColors %>">
					<input type="checkbox" name="default" value="<% = inspecID %>~<% = projectID %>"/></td>
					</tr>
					<% ' Alternate Row Colors
					If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
				End If	
			rsReports2.MoveNext
			Loop
		End If ' END No Results Found
		%>
	</form>
</table>
</body>
</html>

<%
rsReports2.Close
Set rsReports2 = Nothing
connSWPPP.Close
Set connSWPPP = Nothing 
%>
