<%@ Language="VBScript" %>
<%
If _
	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
    Not Session("validErosion") And _
	Not Session("validUser") _
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If
projectID = Request("projID")
projectName = Request("projName")
projectPhase = Request("projPhase")
%><!-- #include file="../admin/connSWPPP.asp" --><%
If Session("validAdmin") Then
	inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, Inspections as i" & _
		" WHERE i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
		" ORDER BY inspecDate DESC"
Else
	inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
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
<html><head><title>SWPPP INSPECTIONS : Report Dates</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../global.css" rel="stylesheet" type="text/css"></head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<table width="100%"><tr><td>
<h1>Please select an inspection date to view the report or&nbsp;&nbsp;
	<button onClick="window.open('reportPrintAll.asp?projID=<%= projectID%>&projName=<%= projectName%>&projPhase=<%= projectPhase %>','','width=800, height=600, location=no, menubar=no, status=no, toolbar=no, scrollbars=yes, resizable=yes')">Print All Reports</button>
	    <br /></div></h1>
<h2><font color="#003399"><% = projectName %>&nbsp;<%= projectPhase %></font></h2>
<% includeItemsFlag = False
rsInspectInfo.MoveFirst()
If rsInspectInfo.EOF Then
	Response.Write("<b><i>Sorry no current " & _
		"data entered at this time.</i></b>")
Else
	inspecID = 0
	Do While Not rsInspectInfo.EOF 
		If rsInspectInfo("released") Then
			If inspecID = 0 Then
				inspecID     = rsInspectInfo("inspecID")
			End If	
            includeItems = rsInspectInfo("includeItems")
			totalItems     = rsInspectInfo("totalItems")
			completedItems = rsInspectInfo("completedItems")
			If includeItems Then
                includeItemsFlag = True
            End If
            If includeItems and Session("seeScoring") and totalItems <> "" Then
                If totalItems <> 0 Then
				    score = " - Report Score: " & FormatNumber((completedItems/totalItems)*100,0) & "% (" & completedItems & "/" & totalItems & ")" 
                Else
                    score = " - Report Score: 100% (" & completedItems & "/" & totalItems & ")" 
                End If
            Else
				score = ""
			End If
			%>
			<br><a href="report.asp?inspecID=<% = rsInspectInfo("inspecID") %>"><% = Trim(rsInspectInfo("inspecDate")) %></a><%=score%>
			<% If Not Session("noImages") Then
	'			imgSQLSELECT = "SELECT COUNT(imageID) FROM Images WHERE inspecID = " & rsInspectInfo("inspecID")
	'			Set rsImages = connSWPPP.execute(imgSQLSELECT)
	'			If rsImages(0)>0 Then
				If rsInspectInfo("ImageCount") > 0 Then%>
					<img src="..\images\smallcamera.gif"><% 
				End If
			End If
		End If
		rsInspectInfo.MoveNext
	Loop
End If ' END No Results Found
'rsImages.Close
'Set rsImages = Nothing
RS0.Close
Set RS0=nothing
rsInspectInfo.Close
Set rsInspectInfo = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
</td><td width="175" valign="top">
<ul>


<!--<li><a href="addOperatorForm.asp?pID=<%= projectID%>" target="_blank">Add Operator Form</a></li>
<li><a href="operatorForm.asp?pID=<%= projectID%>" target="_blank">View Operator Form</a></li>-->
<% If Session("validAdmin") Then %>
    <li><a href="addActionReport.asp?pID=<%= projectID%>" target="_blank">Add Actions Taken</a></li>
    <li><a href="actionReport.asp?pID=<%= projectID%>" target="_blank">View Actions Taken</a></li>
    <li><a href="openActionItems.asp?pID=<%= projectID%>" target="_blank">Open Items</a></li>
    <li><a href="completedActionItems.asp?pID=<%= projectID%>" target="_blank">Completed Items</a></li>
    <li><a href="viewComments.asp?pID=<%=projectID %>" target="_blank">View Item Comments</a></li>
<% Else
    If includeItemsFlag Then
        If Session("seeScoring") Then %>
            <li><a href="openActionItems.asp?pID=<%= projectID%>" target="_blank">Open Items</a></li>
        <% End If %>
        <li><a href="completedActionItems.asp?pID=<%= projectID%>" target="_blank">Completed Items</a></li>
    <% Else
        IF validAct OR Session("validDirector") Then %>
            <li><a href="addActionReport.asp?pID=<%= projectID%>" target="_blank">Add Actions Taken</a></li>
        <% END IF %>
        <li><a href="actionReport.asp?pID=<%= projectID%>" target="_blank">View Actions Taken</a></li>
    <% End If
End If %>

</ul>
</td></tr></table>
</td></tr></table>
</body></html>
