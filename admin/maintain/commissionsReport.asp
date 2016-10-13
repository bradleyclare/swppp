<%@ Language="VBScript" %>
<%
testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
If not(Session("validAdmin") AND InStr(testStr,Session("email"))>0) Then
	Response.Redirect("../default.asp")
End If
startDate=Request("startDate")
endDate=Request("endDate")
If IsNull(endDate) OR (NOT(IsDate(endDate))) Then endDate=DateAdd("d",13-WeekDay(Date()),Date()) End If
If IsDate(endDate) AND IsDate(startDate) Then
	if DateDiff("d", startDate, endDate)<1 then startDate=null end if
End If
If IsNull(startDate) OR (NOT(IsDate(startDate))) Then startDate=DateAdd("d",-6,endDate) End If

%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
SQL1 = "SELECT u.userID, u.lastName, u.firstName, SUM(sum1) commission" &_
	" FROM vCommissionReport vc JOIN Users u ON vc.userID=u.userID" &_
	" WHERE inspecDate Between '"& startDate &"' AND '"& endDate &"'" &_
	" GROUP BY u.userID, u.lastName, u.firstName ORDER BY u.lastName, u.firstName"
Set RS1 = connSWPPP.Execute(SQL1) %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head><title>SWPPP INSPECTIONS : Admin : Commissions</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
<form method="post" action="<% = Request.ServerVariables("script_name") %>">
<table width="100%" border="0">
	<tr><td><br><h1>Commissions</h1></td>
		<td align=left valign=middle>
		From:&nbsp;<INPUT type="text" value="<%= startDate %>" name="startDate" size=10 maxlength=10>
		&nbsp;To:&nbsp;<INPUT type="text" value="<%= endDate %>" name="endDate" size=10 maxlength=10>&nbsp;
		<input type="submit" value="submit"></td>
		</tr></table>
<table width="100%" border="0" cellpadding=0 cellspacing=0>
</form>
	<tr><th align=left>Inspector</th>
		<th align=left>Commission</th></tr>
<%	DO WHILE NOT RS1.EOF 
		if color1="#FFFFFF" then color1="#e5e6e8" else color1="#FFFFFF" end if 
		IF IsNull(RS1("commission")) THEN comm1=0 ELSE comm1=RS1("commission") END IF %>
	<tr	onMouseOver="this.style.backgroundColor='silver'; this.style.cursor='hand';"
		onMouseOut ="this.style.backgroundColor='<%= color1%>'; this.style.cursor='auto';"
		onClick="window.location='commissionsReportDetail.asp?userID=<%= RS1("userID")%>&startDate=<%= startDate%>&endDate=<%= endDate%>'" 
		bgcolor="<%= color1%>"><td><nobr><%= TRIM(RS1("lastName"))%>, <%= TRIM(RS1("firstName"))%></nobr></td>
		<td><%= FormatNumber(comm1,2) %></td></tr>
<%		RS1.MoveNext
	LOOP %>
</table>
</body>
</html><%
RS1.Close
Set connUsers = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>