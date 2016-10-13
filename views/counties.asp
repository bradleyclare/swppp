<%@ Language="VBScript" %><%
If 	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") _	
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("../admin/maintain/loginUser.asp")	
End If
%><!-- #include file="../admin/connSWPPP.asp" --><%
If Session("validAdmin") Then
	cntySQLSELECT = "SELECT DISTINCT projectCounty" & _
		" FROM Inspections ORDER BY projectCounty"
	
Else
	cntySQLSELECT = "SELECT DISTINCT projectCounty" & _
		" FROM Inspections, ProjectsUsers" & _
		" WHERE ProjectsUsers.userID = " & Session("userID") & _
		" AND ProjectsUsers.projectID = Inspections.projectID" & _
		" ORDER BY projectCounty"
'	cntySQLSELECT = "SELECT DISTINCT projectCounty" & _
'		" FROM Inspections as i, ProjectUsers as pu" & _
'		" WHERE pu.userID = " & Session("userID") & _
'		" AND pu.projectID = i.projectID" & _
'		" ORDER BY projectCounty"
End If
Set rsCounty = connSWPPP.Execute(cntySQLSELECT) %>
<html>
<head>
<title>SWPPP INSPECTIONS : Select County</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<TABLE width="100%">
<TR><TD><h1>What county is the project in?</h1></TD>
	<TD><a href="monthlyReportsSum.asp"><h2 align="right">Monthly Summary of Reports<h2></a></TD>
</TR>
</TABLE>
<div class="indent30" align="left"> <%
If rsCounty.EOF Then
	Response.Write("<b><i>Sorry no current " & _
		"data entered at this time.</i></b>")
Else
	Do While Not rsCounty.EOF
		projectCounty = Trim(rsCounty("projectCounty"))
%><a href="projects.asp?cnty=<% = projectCounty %>"><% = projectCounty %></a><br><%
		rsCounty.MoveNext
	Loop
End If ' END No Results Found

rsCounty.Close
Set rsCounty = Nothing

connSWPPP.Close
Set connSWPPP = Nothing%>
</div>
</td></tr></table>	  
</body>
</html>