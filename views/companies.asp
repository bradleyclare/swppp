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

projectCounty = Request("cnty")
%>
<!-- #include file="../admin/connSWPPP.asp" -->
<%
If Session("validAdmin") Then
	coInfoSQLSELECT = "SELECT DISTINCT Companies.companyID, companyName" & _
		" FROM Companies, Inspections" & _
		" WHERE Companies.companyID = Inspections.companyID" & _
		" AND projectCounty = " & "'" & projectCounty & "'" & _
		" ORDER BY companyName"
	
Else
	coInfoSQLSELECT = "SELECT DISTINCT Companies.companyID, companyName" & _
		" FROM Companies, CompanyUsers, Inspections" & _
		" WHERE CompanyUsers.userID = " & Session("userID") & _
		" AND CompanyUsers.companyID = Companies.companyID" & _
		" AND Companies.companyID = Inspections.companyID" & _
		" AND projectCounty = " & "'" & projectCounty & "'" & _
		" ORDER BY companyName"
	
End If

Set rsCoInfo = connSWPPP.Execute(coInfoSQLSELECT)
%>
<html>
<head>
<title>SWPPP Inspections : Select Company</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<h1>What company is working on the project?</h1>
<h2><font color="#003399"><% = projectCounty %></font></h2>
<div class="indent30">
<%
If rsCoInfo.EOF Then
	Response.Write("<b><i>Sorry no current " & _
		"data entered at this time.</i></b>")
Else
	Do While Not rsCoInfo.EOF
		companyID = rsCoInfo("companyID")
		companyName = Trim(rsCoInfo("companyName"))
%>
	<a href="projects.asp?cmpyID=<% = companyID %>&cnty=<% = projectCounty %>"><% = companyName %></a><br>
<%
		rsCoInfo.MoveNext
	Loop
End If ' END No Results Found

rsCoInfo.Close
Set rsCoInfo = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
</div>
</td></tr></table>	  
</body>
</html>
