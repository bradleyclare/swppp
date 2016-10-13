<%@  language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") And Not Session("validDirector") And Not Session("validUser") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If
%><!-- #include virtual="admin/connSWPPP.asp" --><%
IF Request.Form.Count>0 THEN
	SQL0=""
	For Each Item in Request.Form
		IF Request.Form(Item)(1)="on" AND Request.Form(Item).Count=1 THEN
			SQL0=SQL0&" DELETE FROM ProjectsUsers WHERE userID="& Session("userID") &" AND projectID="& Item &" AND rights='email'"
		END IF
		IF Request.Form(Item)(1)="off" AND Request.Form(Item).Count=2 THEN
			SQL0=SQL0&" EXEC sp_InsertPU "& Session("userID") &", "& Item &", 'email'"
		END IF
	Next
'--	Response.Write(SQL0)
	IF LEN(SQL0)>0 THEN connSWPPP.Execute(SQL0) END IF
End If
If Session("validAdmin") Then
	projInfoSQLSELECT = "SELECT DISTINCT p.projectID, p.projectName, p.projectPhase, Case when pu.rights is null then 0 else 1 end as rights " & _
		" FROM Inspections as i inner join Projects p on i.projectid = p.projectid" & _
		"   left join ProjectsUsers pu on p.projectID = pu.projectID and pu.userID = " & Session("userID") &" and pu.rights='email'" & _
		" ORDER BY p.projectName"
Else
	projInfoSQLSELECT = "SELECT DISTINCT p.projectID, p.projectName, p.projectPhase, Case when pu.rights is null then 0 else 1 end as rights " & _
		" FROM Inspections i inner join Projects p on i.projectid = p.projectid" & _
		"   left join ProjectsUsers pu on p.projectID = pu.projectID and pu.userID = " & Session("userID") &" and pu.rights='email'" & _
		" WHERE i.projectID IN" &_
		" (SELECT projectID FROM ProjectsUsers pu WHERE  pu.userID = " & Session("userID") &")" &_
		" ORDER BY p.projectName"
End If
'--Response.Write(projInfoSQLSELECT)
Set rsProjInfo = connSWPPP.Execute(projInfoSQLSELECT)
projectID = rsProjInfo("projectID")
projectName = Trim(rsProjInfo("projectName"))
projectPhase = Trim(rsProjInfo("projectPhase")) %>
<html>
<head>
    <title>SWPPP INSPECTIONS : Select Project</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body>
    <!-- #include virtual="header.inc" -->
    <h1>Inspection Projects</h1>
    <div class="nine columns alpha">
		<table>
			<% IF Session("validUser") OR Session("validAdmin") OR Session("validDirector") THEN %>
			<form action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="POST">
			<% END IF %>
            <tr>
				<% IF Session("validUser") OR Session("validAdmin") OR Session("validDirector") THEN %>
                <th width="20%">Receive Inspections via Email? Check Here and SUBMIT</th>
				<th width="80%">Project Name</th>
				<% Else %>
				<th width="100%">Project Name</th>
				<% END IF %>
            </tr>
			<% If rsProjInfo.EOF Then
				Response.Write("<b><i>Sorry no current data entered at this time.</i></b>")
			Else
				rsProjInfo.MoveFirst
				Do While Not rsProjInfo.EOF
					projectName = Trim(rsProjInfo("projectName"))
					projectPhase= Trim(rsProjInfo("projectPhase")) %>
			<tr>
				<% IF Session("validUser") OR Session("validAdmin") OR Session("validDirector") THEN %>
				<% IF rsProjInfo("rights") = 1 THEN %>
				<td align="right">
					<input type="hidden" name="<%=rsProjInfo("projectID") %>" value="on">
					<input type="checkbox" name="<%=rsProjInfo("projectID") %>" checked>
				</td>
					<% Else %>
				<td align="right">
					<input type="hidden" name="<%=rsProjInfo("projectID") %>" value="off">
					<input type="checkbox" name="<%=rsProjInfo("projectID") %>">
				</td>
					<% END IF %>
				<td align="left">
					<a href="inspections.asp?projID=<% = rsProjInfo("projectID") %>&projName=<% = projectName %>&projPhase=<%= projectPhase %>"><% = projectName %>&nbsp;<%= projectPhase%></a>
				</td>
				<% END IF %>
			</tr>
			<% rsProjInfo.MoveNext
			Loop
			End If%>
			<% IF Session("validUser") OR Session("validAdmin") OR Session("validDirector") THEN %>
			<tr>
				<td align="right">
				<button type="submit">Submit Changes</button>
			</tr>
			</form>
			<% END IF %>
        </table>
    </div>
	<div class="three columns omega">
		<div class="side-link">
			<a href="monthlyReportsSum.asp" target="_blank">Monthly Summary of Inspection Reports</a>
		</div>
		<div class="side-link">
			<a href="reportPrintAllRecent.asp" target="_blank">Print the Latest Inspection Report for Each Project</a>
		</div>
		<div class="side-link">
			<a href="viewAllProjectActions.asp" target="_blank">View Actions Taken for All Projects</a>
		</div>
	</div>
</body>
</html>

<% 
rsProjInfo.Close
Set rsProjInfo = Nothing
connSWPPP.Close
Set connSWPPP = Nothing
%>
