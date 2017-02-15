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

group = Request.QueryString("group")
If isempty(group) Then
    group = "0-9"
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
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<div class="container">
    <h1>View Projects</h1>
    <hr />
    <h3>Select Letter to List Available Projects</h3>
    <ul class="list-inline">
    <li><a href="projects.asp?group=0-9">0-9</a></li>
    <% lastletter = ""
    DO WHILE NOT rsProjInfo.EOF 
        projectName = Trim(rsProjInfo("projectName"))
        letter = LCase(Left(projectName, 1))
        isnumber = isnumeric(letter) or StrComp(letter,"(") = 0
        If not isnumber and StrComp(letter,lastletter) <> 0 Then
            lastletter = letter %>
            <li><a href="projects.asp?group=<% =letter %>"> <% =UCase(letter) %></a></li>
        <% End If		
    rsProjInfo.MoveNext
    LOOP 
    rsProjInfo.MoveFirst%>
    </ul>
    <div class="cleaner"></div>
    <hr />
    <h3><% =Ucase(group) %></h3>
    <table>
	    <tr><td>
		    <table>
			    <tr><th align="left">Project Name</th></tr><%
    If rsProjInfo.EOF Then
	    Response.Write("<b><i>Sorry no current data entered at this time.</i></b>")
    Else
	    Do While Not rsProjInfo.EOF
            letter = LCase(Left(rsProjInfo("projectName"), 1))
            numbergroup = StrComp(group,"0-9") = 0 '0 if equal
            isnumber = isnumeric(letter) or StrComp(letter,"(") = 0
            If LCase(group) = letter or (numbergroup and isnumber) then
		        projectName = Trim(rsProjInfo("projectName"))
		        projectPhase= Trim(rsProjInfo("projectPhase")) %>
			    <tr><td align=left><a href="inspections.asp?projID=<% = rsProjInfo("projectID") %>&projName=<% = projectName %>&projPhase=<%= projectPhase %>"><% = projectName %>&nbsp;<%= projectPhase%></a></td></tr>
            <% End If		
            rsProjInfo.MoveNext
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
				    Monthly Summary of<br />Inspection Reports</a></font></TD></tr>
			    <tr>
				    <td align=center style="border: thin solid #9AB5D1;"
				    onMouseOver="this.style.backgroundColor='#9AB5D1'; this.style.cursor='hand'"
				    onMouseOut="this.style.backgroundColor='transparent'; this.style.cursor='normal'">
				    <font color="black" style="font:normal normal bolder 12px;">
				    <a href="reportPrintAllRecent.asp" target="_blank">
				    Print the Most Recent<br />Inspection Report<br />for Each Project</a></font></td></tr>
                </table>
	        </td></tr>
        </Table>
    </div>
</body>
</html>