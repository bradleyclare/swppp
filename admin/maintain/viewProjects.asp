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
	CASE 1 		orderBy=" Order by phaseNum asc, projectName asc, projectPhase asc"
	CASE 2		orderBy=" Order by initInspecCost asc, projectName asc, projectPhase asc"
	CASE 3		orderBy=" Order by inspecCost asc, projectName asc, projectPhase asc"
	CASE 4		orderBy=" Order by billCycle asc, projectName asc, projectPhase asc"
	CASE else	orderBy=" Order by projectName asc, projectPhase asc"
END SELECT

%> <!-- #include file="../connSWPPP.asp" --> <%
SQL1 = "SELECT * FROM Projects "& orderBy
'-- Response.Write(SQLSELECT & "<br>")
Set RS1 = connSWPPP.Execute(SQL1)
recCount = 0 %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head><title>SWPPP INSPECTIONS : Admin : View Projects</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><br><h1>View Projects</h1></td></tr></table>
<table width="100%" border="0">
	<tr width=50><th align=right><b>Count&nbsp;&nbsp;&nbsp;</b></th>
		<th align=left><b>&nbsp;&nbsp;&nbsp;<a class='head2' href="viewProjects.asp?orderBy=0">Project Name and Phase</a></b></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=1"><b>Comm #</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=2"><b>Init Inspec Cost</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=3"><b>Rec Inspec Cost</b></a></th>
		<th align=center><a class='head2' href="viewProjects.asp?orderBy=4"><b>Bill Cycle</b></a></th>
<%	If RS1.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no Projects.</i></b></td></tr>")
	Else
		altColors="#ffffff"		
		Do While Not RS1.EOF
			recCount = recCount + 1 %>
	<tr align="center" bgcolor="<%= altColors %>"> 
		<td align=right><%= recCount %></td>
		<td align=left><a href="editProjectInfo.asp?id=<%= RS1("projectID") %>">
			<%= Trim(RS1("projectName")) %>&nbsp;<%= Trim(RS1("projectPhase"))%></a></td>
		<td align=center><%= TRIM(RS1("phaseNum"))%></td>
		<td align=center><%= TRIM(RS1("initInspecCost"))%></td>
		<td align=center><%= TRIM(RS1("inspecCost"))%></td>
		<td align=center><%= TRIM(RS1("billCycle"))%></td></tr>
<%			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			RS1.MoveNext
		Loop
	End If ' END No Results Found
RS1.Close
Set connUsers = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
</table>
</body>
</html>