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
projectPhase = Request("projPhase")
%><!-- #include virtual="admin/connSWPPP.asp" --><%
If Session("validAdmin") Then
	inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, Inspections as i" & _
		" WHERE i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
		" ORDER BY inspecDate DESC"
Else
	inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
		" WHERE pu.userID = " & Session("userID") &_
		" AND i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
		" ORDER BY inspecDate DESC"
End If
SQL0="SELECT * FROM ProjectsUsers WHERE "& Session("userID") &" IN (SELECT userID FROM ProjectsUsers WHERE rights in ('action','erosion') AND projectID="& projectID &")"
SET RS0=connSWPPP.execute(SQL0)
'-Response.Write(SQL0 &"<BR>")
validAct=False
IF NOT(RS0.EOF) THEN validAct=True END IF
'Response.Write(inspectInfoSQLSELECT & "<br>")
Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT)
projectName= Trim(rsInspectInfo("projectName"))
projectPhase= Trim(rsInspectInfo("projectPhase")) %>
<html>
<head>
	<title>SWPPP INSPECTIONS : Report Dates</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body>
<!-- #include virtual="header.inc" -->

<h1>Inspection Reports for <br/>
<% = projectName %> <%= projectPhase %></h1>

<div class="nine columns alpha">
	
	<% rsInspectInfo.MoveFirst()
	If rsInspectInfo.EOF Then
		Response.Write("<b><i>Sorry no current data entered at this time.</i></b>")
	Else
		dim cnt
		cnt = 0 %>
	<div class="fl">
    <% Do While Not rsInspectInfo.EOF 
        cnt = cnt + 1
        if cnt Mod 50 = 0 then %>
            </div>
            <div class="fl">
        <% else %>
	        <div style="text-align: right; width: 100px"><a href="report.asp?inspecID=<% = rsInspectInfo("inspecID") %>"><% = Trim(rsInspectInfo("inspecDate")) %></a></div>
            <% IF NOT Session("noImages") THEN
'			imgSQLSELECT = "SELECT COUNT(imageID) FROM Images WHERE inspecID = " & rsInspectInfo("inspecID")
'			Set rsImages = connSWPPP.execute(imgSQLSELECT)
'			If rsImages(0)>0 Then
            If rsInspectInfo("ImageCount") > 0 Then%><img src="..\images\smallcamera.gif"><% End IF
		    END IF
		    rsInspectInfo.MoveNext
        End If
    Loop %>
    </div>
    <div class="cleaner"></div>
<% End If ' END No Results Found %>
</div>
<div class="three columns omega">
<%
'rsImages.Close
'Set rsImages = Nothing
rsInspectInfo.Close
Set rsInspectInfo = Nothing
RS0.Close
Set RS0=nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
<button onClick="window.open('reportPrintAll.asp?projID=<%= projectID%>&projName=<%= projectName%>&projPhase=<%= projectPhase %>','','width=800, height=600, location=no, menubar=no, status=no, toolbar=no, scrollbars=yes, resizable=yes')">Print All Reports</button>
<% IF validAct OR Session("validAdmin") OR Session("validDirector") THEN %>
<div class="side-link">
	<a href="addActionReport.asp?pID=<%= projectID%>" target="_blank">add Actions Taken</a>
</div>
<% END IF %>
<div class="side-link">
	<a href="actionReport.asp?pID=<%= projectID%>" target="_blank">view Actions Taken</a>
</div>
<div class="side-link">
	<a href="addOperatorForm.asp?pID=<%= projectID%>" target="_blank">add Operator Form</a>
</div>
<div class="side-link">
	<a href="operatorForm.asp?pID=<%= projectID%>" target="_blank">view Operator Form</a>
</div>
</body>
</html>