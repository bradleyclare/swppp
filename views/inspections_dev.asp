<%@ Language="VBScript" %>
<%
If _
	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") _
	
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
	
End If

projectID = Request("projID")
projectName = Request("projName")
projectCounty = Request("cnty")
%><!-- #include virtual="admin/connSWPPP.asp" --><%
If Session("validAdmin") Then
	SQL1 = "SELECT DISTINCT inspecID, inspecDate, p.projectName" & _
		" FROM Projects as p, Inspections as i" & _
		" WHERE projectCounty = " & "'" & projectCounty & "'" & _
		" AND p.projectName = " & "'" & projectName & "'" & _ 
		" AND i.projectID=p.projectID" &_
		" ORDER BY inspecDate DESC"	
Else
	SQL1 = "SELECT DISTINCT inspecID, inspecDate, p.projectName" & _
		" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
		" WHERE pu.userID = " & Session("userID") & _
		" AND projectCounty = " & "'" & projectCounty & "'" & _
		" AND p.projectName = " & "'" & projectName & "'" & _
		" AND i.projectID=p.projectID" &_
		" ORDER BY inspecDate DESC"
End If
SQL0="SELECT userID FROM ProjectsUsers WHERE rights='action' AND projectID="& projectID
SET RS0=connSWPPP.execute(SQL0)
validAct=False
IF NOT RS0.eof THEN
	If RS0(0)=Session("userID") THEN validAct=True END IF
END IF
'Response.Write(inspectInfoSQLSELECT & "<br>")
Set RS1 = connSWPPP.Execute(SQL1)

companyName = Trim(RS1("projectName")) %>
<html>
<head>
<title>SWPPP INSPECTIONS : Report Dates</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include virtual="header.inc" -->
<h1>Please select an inspection date to view the report or&nbsp;&nbsp;<Button onClick="PrintAll();">Print All Reports</Button></div></h1>
<h2><font color="#003399"><% = (projectCounty & " : " & projectName) %></font></h2>
<div style="margin-left:10%;">
<%
RS1.MoveFirst()
If RS1.EOF Then
	Response.Write("<b><i>Sorry no current " & _
		"data entered at this time.</i></b>")
Else
	Do While Not RS1.EOF
		inspecID = RS1("inspecID")
		inspecDate = Trim(RS1("inspecDate")) %>
	<a href="report.asp?inspecID=<% = inspecID %>"><% = inspecDate %></a><br>
<%		RS1.MoveNext
	Loop
End If ' END No Results Found
RS1.Close
Set RS1 = Nothing
RS0.Close
SET RS0=nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
<p class="indent30"><a href="actionReport.asp?pID=<%= projectID%>" target="_blank">click to view Actions Taken Report</a></p>
<% IF validAct THEN%>
	<p class="indent30"><a href="addActionReport.asp?pID=<%= projectID%>" target="_blank">click to Add and Actions Taken Event</a></p>
<% END IF %>
</div>
</td></tr></table>	  
</body>
</html>