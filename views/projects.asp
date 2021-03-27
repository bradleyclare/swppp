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

%><!-- #include file="../admin/connSWPPP.asp" --><%
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
Set rsProjInfo = connSWPPP.Execute(projInfoSQLSELECT) %>
<html>
<head>
<title>SWPPP INSPECTIONS : Select Project</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<div class="indent30">
<table>
	<tr><td>
		<table>
			<tr><th align="left">Project Name</th></tr><%
If rsProjInfo.EOF Then
	Response.Write("<b><i>Sorry no current data entered at this time.</i></b>")
Else
    projectID = rsProjInfo("projectID")
    projectName = Trim(rsProjInfo("projectName"))
    projectPhase = Trim(rsProjInfo("projectPhase")) 
	rsProjInfo.MoveFirst
	Do While Not rsProjInfo.EOF
		projectName = Trim(rsProjInfo("projectName"))
		projectPhase= Trim(rsProjInfo("projectPhase")) %>
			<tr><td align=left><a href="inspections.asp?projID=<% = rsProjInfo("projectID") %>&projName=<% = projectName %>&projPhase=<%= projectPhase %>"><% = projectName %>&nbsp;<%= projectPhase%></a></td></tr>
<%		rsProjInfo.MoveNext
	Loop
End If ' END No Results Found
rsProjInfo.Close
Set rsProjInfo = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
    </table></td>
		<td valign="top">
         <table align="left">
			<tr><td>&nbsp;</td></tr>
			<tr>
				<TD align=center style="border: thin solid #9AB5D1;"
				onMouseOver="this.style.backgroundColor='#9AB5D1'; this.style.cursor='hand'"
				onMouseOut="this.style.backgroundColor='transparent'; this.style.cursor='normal'">
				<font color="black" style="font:normal normal bolder 12px;">
				<a href="monthlyReportsSum.asp" target="_blank">
				monthly summary of<br />inspection reports</a></font></TD></tr>
			<tr>
				<td align=center style="border: thin solid #9AB5D1;"
				onMouseOver="this.style.backgroundColor='#9AB5D1'; this.style.cursor='hand'"
				onMouseOut="this.style.backgroundColor='transparent'; this.style.cursor='normal'">
				<font color="black" style="font:normal normal bolder 12px;">
				<a href="reportPrintAllRecent.asp" target="_blank">
				print the most recent<br />inspection report<br />for each project</a></font></td></tr>
			<tr>
				<td align=center style="border: thin solid #9AB5D1;"
				onMouseOver="this.style.backgroundColor='#9AB5D1'; this.style.cursor='hand'"
				onMouseOut="this.style.backgroundColor='transparent'; this.style.cursor='normal'">
				<font color="black" style="font:normal normal bolder 12px;">
				<% If Session("validAdmin") Then %>
					<a href="../admin/maintain/recentComments.asp" target="_blank">
				<% Else %>
					<a href="../admin/maintain/recentCommentsUser.asp" target="_blank">
				<% End If %>
				notes</a></font></td></tr>
			</table></td></tr>
    </Table>
</div>
</body>
</html>