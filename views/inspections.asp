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

currentDate = date()
'Response.Write(currentDate)
If Request.Form.Count > 0 Then
	SQLI="SELECT inspecID FROM Inspections WHERE horton=1 AND projectID="& projectID
	SET RSI=connSWPPP.execute(SQLI)
	Do While Not RSI.EOF
		inspecID = RSI("inspecID")
		userID = Session("userID")
		if Request("approval_V:" & inspecID ) = "on" Then
			'log the report approval in the database, check if it exists
			SQLHA="SELECT * FROM HortonApprovals WHERE LD=0 and inspecID="& inspecID
			SET RSHA=connSWPPP.execute(SQLHA)
			If RSHA.EOF Then
				'add new entry
				SQLHU = "INSERT into HortonApprovals (inspecID, userID, date, LD) VALUES (" & inspecID & ", " & userID & ", '" & date() & "', 0)"
			Else
				'update the entry
				SQLHU = "UPDATE HortonApprovals set userID=" & userID & ", date=" & currentDate & " WHERE inspecID=" & inspecID
			End If
			'Response.Write(SQLHU)
			connSWPPP.Execute(SQLHU)
			'check the date
			SET RSC=connSWPPP.execute("SELECT TOP 1 * FROM HortonApprovals ORDER BY id DESC")
			If RSC("date") = "1/1/1900" Then
				SQLHU = "UPDATE HortonApprovals set date='" & date() & "' WHERE id=" & RSC("id")
				connSWPPP.Execute(SQLHU)
			End If
		ElseIf Request("approval_LD:" & inspecID ) = "on" Then
			'log the report approval in the database, check if it exists
			SQLHA="SELECT * FROM HortonApprovals WHERE LD=1 and inspecID="& inspecID
			SET RSHA=connSWPPP.execute(SQLHA)
			If RSHA.EOF Then
				'add new entry
				SQLHU = "INSERT into HortonApprovals (inspecID, userID, date, LD) VALUES (" & inspecID & ", " & userID & ", '" & date() & "', 1)"
			Else
				'update the entry
				SQLHU = "UPDATE HortonApprovals set userID=" & userID & ", date=" & currentDate & " WHERE inspecID=" & inspecID
			End If
			'Response.Write(SQLHU)
			connSWPPP.Execute(SQLHU)
			'check the date
			SET RSC=connSWPPP.execute("SELECT TOP 1 * FROM HortonApprovals ORDER BY id DESC")
			If RSC("date") = "1/1/1900" Then
				SQLHU = "UPDATE HortonApprovals set date='" & date() & "' WHERE id=" & RSC("id")
				connSWPPP.Execute(SQLHU)
			End If
		End If		
		RSI.MoveNext
	Loop
End If

If Session("validAdmin") Then
	inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, horton, hortonSignV, hortonSignLD, vscr, ldscr, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, Inspections as i" & _
		" WHERE i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
		" ORDER BY inspecDate DESC"
Else
	inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, horton, hortonSignV, hortonSignLD, vscr, ldscr, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
		" WHERE pu.userID = " & Session("userID") &_
		" AND i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
      " ORDER BY inspecDate DESC"
End If
'Response.Write(inspectInfoSQLSELECT & "<br>")
Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT)
projectName= Trim(rsInspectInfo("projectName"))
projectPhase= Trim(rsInspectInfo("projectPhase"))
'SQL0="SELECT * FROM ProjectsUsers WHERE "& Session("userID") &" IN (SELECT userID FROM ProjectsUsers WHERE rights in ('action','erosion') AND projectID="& projectID &")"
'SET RS0=connSWPPP.execute(SQL0)
'Response.Write(SQL0 &"<BR>")
'validAct=False
'IF NOT(RS0.EOF) THEN validAct=True END IF
SQL1="SELECT inspecID FROM Inspections WHERE horton=1 AND projectID="& projectID
SET RS1=connSWPPP.execute(SQL1)
hortonFlag=False
completePast="completed"
if NOT(RS1.EOF) THEN 
	hortonFlag=True 
	completePast="closed"
END IF

SQL1="SELECT inspecID FROM Inspections WHERE hortonSignV=1 AND projectID="& projectID
SET RS1=connSWPPP.execute(SQL1)
hortonSignV=False
if NOT(RS1.EOF) THEN hortonSignV=True END IF

SQL1="SELECT inspecID FROM Inspections WHERE hortonSignLD=1 AND projectID="& projectID
SET RS1=connSWPPP.execute(SQL1)
hortonSignLD=False
if NOT(RS1.EOF) THEN hortonSignLD=True END IF

SQL1="SELECT * FROM ProjectsUsers WHERE userID="& Session("userID") & " AND projectID="& projectID &" AND rights='vscr'"
SET RS1=connSWPPP.execute(SQL1)
hortonUserV=False
if NOT(RS1.EOF) THEN hortonUserV=True END IF

SQL1="SELECT * FROM ProjectsUsers WHERE userID="& Session("userID") & " AND projectID="& projectID &" AND rights='ldscr'"
SET RS1=connSWPPP.execute(SQL1)
hortonUserLD=False
if NOT(RS1.EOF) THEN hortonUserLD=True END IF
'Response.Write("projectID: " & projectID & ", vscr: " & hortonSignV & ", ldscr: " & hortonSignLD & "<br/>")
 %>

<html>
<head>
	<title>SWPPP INSPECTIONS : Report Dates</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?projID=<%=projectID%>&projName=<%=projectName%>&projPhase=<%=projectPhase%>%>" onsubmit="return isReady(this)";>
<table width="100%"><tr><td>
<h1><font color="#003399"><% = projectName %>&nbsp;<%= projectPhase %></font></h1>
<table>
<tr><td><button onClick="window.open('reportPrintAll.asp?projID=<%= projectID%>&projName=<%= projectName%>&projPhase=<%= projectPhase %>','','width=800, height=600, location=no, menubar=no, status=no, toolbar=no, scrollbars=yes, resizable=yes')">print all reports</button></td>
<td>
<% If hortonFlag AND (hortonUserV or hortonUserLD or Session("validAdmin") or Session("validDirector")) Then %>
<input type="submit" value="sign off on reports" name="approve_reports"></input>
<% End If %>
</td></tr>
</table>
<br />
<table>

<tr><th>report date</th><th>report</th><th>site map</th>
<%  'set proper header for user rights and permissions
If Session("seeScoring") Then %>
	<th>report score</th><th>items</th>
<% End If 
If hortonFlag Then
	If hortonSignV and (Session("validAdmin") or Session("validDirector") or hortonUserV) Then %>	
			<th>VSCR sign off</th><th>VSCR</th><th>VSCR date</th>
	<% End If
   If hortonSignLD and (Session("validAdmin") or Session("validDirector") or hortonUserLD) Then %>	
			<th>LDSCR sign off</th><th>LDSCR</th><th>LDSCR date</th>
	<% End If
End If %>
</tr>

<% includeItemsFlag = False
firstInspecID = 0
rsInspectInfo.MoveFirst()
If rsInspectInfo.EOF Then
	Response.Write("<b><i>Sorry no current " & _
		"data entered at this time.</i></b>")
Else
	inspecID = 0
	Do While Not rsInspectInfo.EOF
	   'Response.Write(inspecID & "<br/>") 
		If inspecID = 0 Then
			firstInspecID     = rsInspectInfo("inspecID")
		End If
		If rsInspectInfo("released") Then
         inspecID = rsInspectInfo("inspecID")
         includeItems = rsInspectInfo("includeItems")
			totalItems     = rsInspectInfo("totalItems")
			completedItems = rsInspectInfo("completedItems")
			hortonFlag = rsInspectInfo("horton")
			hortonSignV = rsInspectInfo("hortonSignV")
			hortonSignLD = rsInspectInfo("hortonSignLD")
			'Response.Write("inspecID: " & inspecID & ", vscr: " & hortonSignV & ", ldscr: " & hortonSignLD & "<br/>")
			If includeItems Then
            includeItemsFlag = True
         End If
         If includeItems and Session("seeScoring") and totalItems <> "" Then
             If totalItems <> 0 Then
                percentage = FormatNumber((completedItems/totalItems)*100,0) & "%"
             Else
                percentage = "100%"
             End If
             stats = "(" & completedItems & "/" & totalItems & ")"
         Else
			   percentage = ""
            stats = ""
			End If
			%>
			<tr><td align="right"><% = Trim(rsInspectInfo("inspecDate")) %></td>
            <td align="center"><a href="reportPrint.asp?inspecID=<% = inspecID %>" target="_blank">report</a></td>
            <td align="center"><a href="viewSitemap.asp?inspecID=<% = inspecID %>" target="_blank">site map</a></td>
			<% If Session("seeScoring") Then %>
				<td align="center"><%=percentage%></td>
            <td align="center"><%=stats%></td>
			<% End If 
			If hortonFlag Then 
				'check for approval status 
				If hortonSignV and (Session("validAdmin") or Session("validDirector") or hortonUserV) Then
					SQLA="SELECT * FROM HortonApprovals WHERE LD=0 and inspecID="& inspecID
					SET RSA=connSWPPP.execute(SQLA)
					If RSA.EOF Then 
						hortonStatus = False
						hortonApprovalUser = ""
						hortonApprovalDate = ""
					Else
						hortonStatus = True
						SQLU="SELECT firstName, lastName FROM Users WHERE userID="& RSA("userID")
						SET RSU=connSWPPP.execute(SQLU)
						'Response.Write("userID:" & RSA("userID"))
						If not RSU.EOF Then
							hortonApprovalUser = RSU("firstName") & " " & RSU("lastName")
						Else
							hortonApprovalUser = "Unknown"
						End If
						hortonApprovalDate = RSA("date")
					End If %>
					<td align="center">
					<% If hortonStatus Then %>
						x
					<% Else %>
						<input type="checkbox" name="approval_V:<%=inspecID%>"></input>
					<% End If %>
					</td>
					<td align="center"><%=hortonApprovalUser%></td>
					<td align="center"><%=hortonApprovalDate%></td>
				<% End If
				If hortonSignLD and (Session("validAdmin") or Session("validDirector") or hortonUserLD) Then
					SQLA="SELECT * FROM HortonApprovals WHERE LD=1 and inspecID="& inspecID
					SET RSA=connSWPPP.execute(SQLA)
					If RSA.EOF Then 
						hortonStatus = False
						hortonApprovalUser = ""
						hortonApprovalDate = ""
					Else
						hortonStatus = True
						SQLU="SELECT firstName, lastName FROM Users WHERE userID="& RSA("userID")
						SET RSU=connSWPPP.execute(SQLU)
						If not RSU.EOF Then
							hortonApprovalUser = RSU("firstName") & " " & RSU("lastName")
						Else
							hortonApprovalUser = "Unknown"
						End If
						hortonApprovalDate = RSA("date")
					End If %>
					<td align="center">
					<% If hortonStatus Then %>
						x
					<% Else %>
						<input type="checkbox" name="approval_LD:<%=inspecID%>"></input>
					<% End If %>
					</td>
					<td align="center"><%=hortonApprovalUser%></td>
					<td align="center"><%=hortonApprovalDate%></td>
				<% End If
			End If %>
			</tr>
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
'RS0.Close
'Set RS0=nothing
rsInspectInfo.Close
Set rsInspectInfo = Nothing %>
</table>
</td>   
<td width="175" valign="top">
<h5>project management</h5>
<ul>
<!--<li><a href="addOperatorForm.asp?pID=<%= projectID%>" target="_blank">add operator Form</a></li>
<li><a href="operatorForm.asp?pID=<%= projectID%>" target="_blank">view operator form</a></li>-->
<% If Session("validAdmin") Then %>
    <li><a href="addActionReport.asp?pID=<%= projectID%>" target="_blank">add actions taken</a></li>
    <li><a href="actionReport.asp?pID=<%= projectID%>" target="_blank">view actions taken</a></li>
    <li><a href="openActionItems.asp?pID=<%= projectID%>" target="_blank">open items</a></li>
    <li><a href="completedActionItems.asp?pID=<%= projectID%>" target="_blank"><%=completePast%> items</a></li>
<% Else %>
   <li><a href="viewComments.asp?pID=<%=projectID %>" target="_blank">view item notes</a></li>
    <% If includeItemsFlag Then
        If Session("seeScoring") Then %>
            <li><a href="openActionItems.asp?pID=<%= projectID%>" target="_blank">open items</a></li>
        <% End If %>
        <li><a href="completedActionItems.asp?pID=<%= projectID%>" target="_blank"><%=completePast%> items</a></li>
    <% End If
End If %>
</ul>
<% If Not Session("validErosion") Then %>
<h5>project documents</h5>
<% End If %>
<ul>
   <% SQL2="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=-1 ORDER BY oitSortByVal asc"
   SET RS2=connSWPPP.execute(SQL2)
   DO WHILE NOT RS2.EOF
	   dirName=Trim(RS2("oitName"))
	   fileDesc= TRIM(RS2("oitDesc"))
	   SQLA="sp_oImagesByType "& firstInspecID &",'"& RS2("oitID") &"'"
	   SET RSA=connSWPPP.Execute(SQLA)
	   'Response.Write(SQLA)
	   cnt1=1
	   curOITDesc=""
 	   DO WHILE NOT(RSA.EOF)
		   thisFileDesc=fileDesc
		   if curOITDesc=fileDesc then
			   cnt1=cnt1+1
		   else
			   cnt1=1
			   curOITDesc=fileDesc
		   end if
		   if (cnt1>1) then thisFileDesc=thisFileDesc &" "& cnt1 end if
		   IF dirName <> "sitemap" THEN
			   If Not Session("validErosion") Then %>
      	   <li><a href="<% = "../images/"& dirName &"/"& RSa("oImageFileName") %>" target="_blank"><%= thisFileDesc%></a></li>
   			<% End If
		   END IF
		   RSA.MoveNext
	   LOOP
	   RS2.MoveNext
   LOOP 
connSWPPP.Close
Set connSWPPP = Nothing %>
</ul>
</td></tr></table>
</td></tr></table>
</form>
</body></html>
