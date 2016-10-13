<%@ Language="VBScript" %>
<%
testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
If not(Session("validAdmin") AND InStr(testStr,Session("email"))>0) Then
	Response.Redirect("../default.asp")
End If
userID=Request("userID")
startDate=Request("startDate")
endDate=Request("endDate")
If IsNull(userID) OR NOT(IsNumeric(userID)) Then Response.Redirect("commissionsReport.asp") End IF
If IsNull(endDate) OR (NOT(IsDate(endDate))) Then endDate=DateAdd("d",13-WeekDay(Date()),Date()) End If
If IsDate(endDate) AND IsDate(startDate) Then
	if DateDiff("d", startDate, endDate)<1 then startDate=null end if
End If
If IsNull(startDate) OR (NOT(IsDate(startDate))) Then startDate=DateAdd("d",-6,endDate) End If

%> <!-- #include file="../connSWPPP.asp" --> <%
SQL0="SELECT * FROM Users WHERE userID="& userID
SQL1 = "SELECT p.projectID, p.projectName, p.projectPhase, vc.inspecDate, IsNull(vc.sum1,0.00) as sum1 " &_
	" FROM vCommissionReport vc, Projects p WHERE p.projectID = vc.projectID" &_
	" AND userID="& userID &" AND vc.inspecDate Between '"& startDate &"' AND '"& endDate &"'" &_
	" ORDER BY p.projectName, p.projectPhase, inspecDate DESC"
Set RS0 = connSWPPP.Execute(SQL0) 
Set RS1 = connSWPPP.Execute(SQL1) %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html><head><title>SWPPP INSPECTIONS : Admin : Commissions</title>
	<link rel="stylesheet" href="../../global.css" type="text/css"></head>
<table width="100%" border="0">
	<tr><td><h1>Commissions Detail for <%= TRIM(RS0("firstName"))%> <%= TRIM(RS0("lastName"))%>
		<br>From: <%= startDate %> To: <%= endDate %></h1></td></tr></table>
<table width="100%" border="0" cellpadding=0 cellspacing=0>
	<tr><th align=left>Project | Phase</th>
		<th align=left>Inspection Date</th>
		<th align=left>Commission Amount</th></tr><%
	curProjID=-1
	DO WHILE NOT RS1.EOF 
		if color1="#FFFFFF" then color1="#e5e6e8" else color1="#FFFFFF" end if %>
	<tr bgcolor="<%= color1%>"><%
		if curProjID<>RS1("projectID") then 
			curProjID=RS1("projectID") %>
			<td><%= Trim(RS1("projectName")) %> <%= Trim(RS1("projectPhase")) %></td><%
		else %>
			<td>&nbsp;</td><%
		end if %>
		<td><%= Trim(RS1("inspecDate")) %></td>
			<td><%= FormatNumber(RS1("sum1"),2) %></td></tr>
<%		RS1.MoveNext
	LOOP %>
</table>
</body>
</html><%
RS1.Close
Set connUsers = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>