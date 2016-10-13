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

projectCounty = Request("cnty")
%><!-- #include virtual="admin/connSWPPP.asp" --><%
If Session("validAdmin") Then
	projInfoSQLSELECT = "SELECT DISTINCT i.projectID, p.projectName, i.projectPhase" & _
		" FROM Inspections as i inner join Projects as p on i.projectid = p.projectid WHERE i.projectCounty = " & "'" & projectCounty & "'" &_
		" ORDER BY i.projectName"
Else
	projInfoSQLSELECT = "SELECT DISTINCT i.projectID, p.projectName, i.projectPhase" & _
		" FROM Inspections i inner join Projects as p on i.projectid = p.projectid WHERE i.projectCounty = " & "'" & projectCounty & "'" & _
		" AND i.projectID IN (SELECT projectID FROM ProjectsUsers pu WHERE  pu.userID = " & Session("userID") & _
		") ORDER BY i.projectName"
End If
'--Response.Write(projInfoSQLSELECT)
Set rsProjInfo = connSWPPP.Execute(projInfoSQLSELECT)
projectID = rsProjInfo("projectID")
projectName = Trim(rsProjInfo("projectName")) %>
<html>
<head>
<title>SWPPP INSPECTIONS : Select Project</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include virtual="header2.inc" -->
<h1>What is the project name?</h1>
<% '----- (projectCounty & " : " & companyName) --- original h2 code %>
<table width="100%" cellpadding="0" cellspacing="0" background="0">
	<tr><td align=left><h2><font color="#003399"><%= projectCounty %></font></h2></td>
		<td align=right valign="top"><a href="reportPrintAllRecent.asp?cnty=<%=projectCounty%>" target="_blank">
			<h2>Print the Latest Inspection Report for Each Project</h2></a></td></tr>
</table>
<div class="indent30">
<table><%
IF Session("validUser") THEN %><th nowrap>Email Reports</th><% END IF %>
	<th align="left">Project Name</th>
<%
If rsProjInfo.EOF Then
	Response.Write("<b><i>Sorry no current " & _
		"data entered at this time.</i></b>")
Else
	Do While Not rsProjInfo.EOF
		projectName = Trim(rsProjInfo("projectName"))
		projectPhase= Trim(rsProjInfo("projectPhase")) %>
	<tr>
<%		IF Session("validUser") THEN %>
<%			SQL1="SELECT emailReport FROM ProjectsUsers WHERE rights='user' AND userID="& Session("userID") &" AND projectID="& rsProjInfo("projectID")
			SET RS1=connSWPPP.Execute(SQL1)
 			IF NOT(RS1.BOF AND RS1.EOF) THEN
				If RS1("emailReport") THEN %>
		<td align=right><input type="checkbox" name="projID<%=rsProjInfo("projectID") %>" checked></td>
<% 				End If %>
<%			Else %>
		<td><input type="checkbox" name="projID<%=rsProjInfo("projectID") %>"></td>
<%			End If
		END IF %><td align=left><a href="inspections.asp?projID=<% = rsProjInfo("projectID") %>&projName=<% = projectName %>&cnty=<% = projectCounty %>"><% = projectName %>&nbsp;<%= projectPhase%></a></td></tr>
<%		rsProjInfo.MoveNext
	Loop
End If ' END No Results Found
rsProjInfo.Close
Set rsProjInfo = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
</Table>
</div>
</td></tr></table>
</body>
</html>