<%@ Language="VBScript" %>
<%
testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
If not(Session("validAdmin") AND InStr(testStr,Session("email"))>0) Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

If Session("validDirector") then
	Response.Redirect("viewUsersDir.asp")
end if

recordOrd = Request("orderBy")
SELECT CASE recordOrd
	CASE 1 		orderBy=" Order by active asc, projectName asc, projectPhase asc"
	CASE 2		orderBy=" Order by initInspecCost asc, projectName asc, projectPhase asc"
	CASE 3		orderBy=" Order by inspecCost asc, projectName asc, projectPhase asc"
	CASE 4		orderBy=" Order by billCycle asc, projectName asc, projectPhase asc"
    CASE 5      orderBy=" Order by collectionName asc, projectName asc, projectPhase asc"
	CASE else	orderBy=" Order by projectName asc, projectPhase asc"
END SELECT

%> <!-- #include file="../connSWPPP.asp" --> <%
recCount = 0 
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head><title>SWPPP INSPECTIONS : Admin : View Projects</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<%
SQL1 = "SELECT * FROM Projects WHERE active=1 "& orderBy
'-- Response.Write(SQL1 & "<br>")
Set RS1 = connSWPPP.Execute(SQL1)
%>
<h1>View Projects</h1>
<h2>Active Projects</h2>
<table width="100%" border="0">
	<tr width=50><th align=right><b>Count&nbsp;&nbsp;&nbsp;</b></th>
		<th align=left><b>&nbsp;&nbsp;&nbsp;<a class='head2' href="viewProjects.asp?orderBy=0">Project Name and Phase</a></b></th>
      <th align=center><a class='head2'><b>Manage Rights</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=1"><b>Active</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=2"><b>Init Inspec Cost</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=3"><b>Rec Inspec Cost</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=4"><b>Bill Cycle</b></a></th>
        <th align=center><a class='head2' href="viewProjects.asp?orderBy=5"><b>Collection Name</b></a></th>
<%	If RS1.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no Projects.</i></b></td></tr>")
	Else
		activeColor="#ffffff"
		inactiveColor="#e5e6e8"	
		Do While Not RS1.EOF
			recCount = recCount + 1 
			projID = RS1("projectID")
			projName = Trim(RS1("projectName")) & " " & Trim(RS1("projectPhase"))
			active = Trim(RS1("active"))
			initInspecCost = TRIM(RS1("initInspecCost"))
			inspecCost = TRIM(RS1("inspecCost"))
			billCycle = TRIM(RS1("billCycle"))
			collectionName = TRIM(RS1("collectionName"))
			If active Then
				color = activeColor
			Else
				color = inactiveColor
			End If
			%>
			<tr align="center" bgcolor="<%= color %>"> 
				<td align=right><%= recCount %></td>
				<td align=left><a href="editProjectInfo.asp?id=<%= projID %>">
				<%= projName %></a></td>
				<td><a href="editUsersByProject.asp?pID=<%= projID %>">rights</a></td>
				<td align=center><%= active %></td>
				<td align=center><%= initInspecCost %></td>
				<td align=center><%= inspecCost %></td>
				<td align=center><%= billCycle %></td>
				<td align=center><%= collectionName %></td>
     		</tr>
			<% RS1.MoveNext
		Loop
	End If ' END No Results Found
%> </table>
<%
SQL2 = "SELECT * FROM Projects WHERE active=0 "& orderBy
'-- Response.Write(SQL2 & "<br>")
Set RS2 = connSWPPP.Execute(SQL2)
%>
<h2>Inactive Projects</h2>
<table width="100%" border="0">
	<tr width=50><th align=right><b>Count&nbsp;&nbsp;&nbsp;</b></th>
		<th align=left><b>&nbsp;&nbsp;&nbsp;<a class='head2' href="viewProjects.asp?orderBy=0">Project Name and Phase</a></b></th>
      <th align=center><a class='head2'><b>Manage Rights</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=1"><b>Active</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=2"><b>Init Inspec Cost</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=3"><b>Rec Inspec Cost</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=4"><b>Bill Cycle</b></a></th>
        <th align=center><a class='head2' href="viewProjects.asp?orderBy=5"><b>Collection Name</b></a></th>
<%	If RS2.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no Projects.</i></b></td></tr>")
	Else
		activeColor="#ffffff"
		inactiveColor="#e5e6e8"	
		Do While Not RS2.EOF
			recCount = recCount + 1 
			projID = RS2("projectID")
			projName = Trim(RS2("projectName")) & " " & Trim(RS2("projectPhase"))
			active = Trim(RS2("active"))
			initInspecCost = TRIM(RS2("initInspecCost"))
			inspecCost = TRIM(RS2("inspecCost"))
			billCycle = TRIM(RS2("billCycle"))
			collectionName = TRIM(RS2("collectionName"))
			If active Then
				color = activeColor
			Else
				color = inactiveColor
			End If
			%>
			<tr align="center" bgcolor="<%= color %>"> 
				<td align=right><%= recCount %></td>
				<td align=left><a href="editProjectInfo.asp?id=<%= projID %>">
				<%= projName %></a></td>
				<td><a href="editUsersByProject.asp?pID=<%= projID %>">rights</a></td>
				<td align=center><%= active %></td>
				<td align=center><%= initInspecCost %></td>
				<td align=center><%= inspecCost %></td>
				<td align=center><%= billCycle %></td>
				<td align=center><%= collectionName %></td>
     		</tr>
			<% RS2.MoveNext
		Loop
	End If ' END No Results Found
%> </table>

<% RS1.Close
RS2.Close
Set connUsers = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
</body>
</html>