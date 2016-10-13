<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

recordOrd = Request("orderBy")
If recordOrd = "" Then recordOrd = "inspecDate" End If
%>
<!-- #include virtual="admin/connSWPPP.asp" -->
<html>
	<head>
	<title>SWPPP INSPECTIONS : Select New or Default Inspection Report</title>
	<link rel="stylesheet" type="text/css" href="../../global.css"/>
</head>
<body>
	<!-- #include virtual="admin/adminHeader2.inc" -->
	
	<div class="six columns alpha">
		<h3>Select New or Default Inspection Report</h3>
	</div>
	<div class="six columns omega">
		<form method="post" name="form_new_report" action="newReport.asp">
			<input type="submit" value="New Inspection Report"/>
		</form>
	</div>
	<p>Or, choose an inspection report using default company and inspection data.</p>
	<form method="post" name="form1" action="addReport.asp">
	<table width="100%" border="0">
		<tr><th width="15%"><a href="<%= Request.ServerVariables("script_name") %>?orderBy=inspecDate"><b>Date</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=projectName"><b>Company</b></a></th>
		<th width="15%"><input type=submit value="Create Inspections" /></th>
		</tr>
<%	SQL0=" SELECT i.inspecID, i.inspecDate, i.projectName, i.projectPhase, p.projectID "&_
        " FROM Inspections i inner join Projects p on i.projectid = p.projectid inner join (" &_
        "   Select projectID, MAX(inspecDate) inspecDate From Inspections Group By projectID) as i2 " &_
        "       on i.projectID = i2.projectID and i.inspecDate = i2.inspecDate and DateDiff(mm, i.inspecDate, GetDate()) < 3"
    IF session("validInspector") AND NOT(Session("validAdmin")) THEN SQL0 = SQL0 & " WHERE i.userID='" & Session("userID") &"'"
    SQL0 = SQL0 &" ORDER BY i.projectName asc"
	SET rsReports2=connSWPPP.Execute(SQL0)
    If rsReports2.EOF Then
		Response.Write("<tr><td colspan='3' align='center'><b><i>Sorry " & _
			"no current data entered at this time.</i></b></td></tr>")
	Else
		altColors = "#e5e6e8"
		Do While Not rsReports2.EOF
			inspecID = rsReports2("inspecID")
			inspecDate = Trim(rsReports2("inspecDate"))
			projectName = Trim(rsReports2("projectName"))
			projectPhase = Trim(rsReports2("projectPhase"))
			projectID = rsReports2("projectID")%>
		<tr nowrap><td align="center" bgcolor="<% = altColors %>"><% = inspecDate %></td>
			<td align="center" bgcolor="<% = altColors %>"><% = projectName %>&nbsp;<% = projectPhase %></td>
			<td align="center" bgcolor="<% = altColors %>">
				<input type="checkbox" name="default" value="<% = inspecID %>~<% = projectID %>"/></td>
		</tr>
<% 			' Alternate Row Colors
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			rsReports2.MoveNext
		Loop
	End If ' END No Results Found
rsReports2.Close
Set rsReports2 = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
		</table>
		</div>
	</form>
</body>
</html>
