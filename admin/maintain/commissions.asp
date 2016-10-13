<%@ Language="VBScript" %>
<%
testStr="dwims@swppp.com:jwright@swppp.com"
If not(Session("validAdmin") AND InStr(testStr,Session("email"))>0) Then
	Response.Redirect("../default.asp")
End If
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%

IF Request.Form.Count>0 THEN
	SQL0=""
	FOR EACH Item IN Request.Form
		arrID=SPLIT(Item,":")
		arrPhase=SPLIT(Request(Item),",")
		SQL0=SQL0 &" EXEC sp_UpdateCommissions "& arrID(0) &", '"& arrID(1) &"', "& arrPhase(0) &", "& arrPhase(1) &", "& arrPhase(2) &", "& arrPhase(3) &", "& arrPhase(4) 
	NEXT
	connSWPPP.execute(SQL0)
END IF

SQLSELECT = "SELECT DISTINCT c.userID, c.phase1, c.phase2, c.phase3, c.phase4, c.phase5, u.lastName, u.firstName, p.projectName " &_
    " FROM ProjectsUsers pu JOIN Users u on pu.userID=u.userID JOIN Projects p ON pu.projectID=p.projectID JOIN Commissions c ON u.userID = c.userID and p.projectID = c.projectID " &_
    " WHERE u.userID IN (SELECT DISTINCT userID FROM ProjectsUsers WHERE rights='inspector') AND pu.rights='inspector' " &_
    " ORDER BY lastName, firstName, p.projectName"
'-- Response.Write(SQLSELECT & "<br>")
Set RS1 = connSWPPP.Execute(SQLSELECT)
recCount = 0 %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head><title>SWPPP INSPECTIONS : Admin : Commissions</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><br><h1>Commissions</h1></td>
	<td align=right valign=middle><a href="commissionsReport.asp">...goto Commissions Report</a></tr></table>
<table width="100%" border="0" cellpadding=0 cellspacing=0>
	<FORM action="<% = Request.ServerVariables("script_name") %>" method="POST">
	<tr><th align=left>Inspector</th>
		<th align=left>Project</th>
		<th>Comm1</th>
		<th>Comm2</th>
		<th>Comm3</th>
		<th>Comm4</th>
		<th>Comm5</th></tr>
<%	currFullName=""
	DO WHILE NOT RS1.EOF 
		if color1="#FFFFFF" then color1="#e5e6e8" else color1="#FFFFFF" end if 
		fullName= Trim(RS1("lastName")) &", "& Trim(RS1("firstName")) %>
<%		IF currFullName<>fullName THEN
			color1="" 
			currFullName=fullName %>
	<tr><td colspan=7>&nbsp;</td></tr><tr bgcolor="<%= color1%>"><td><nobr><%= fullName%></nobr></td>
<%		ELSE %>
	<tr bgcolor="<%= color1%>"><td>&nbsp;</td>
<% 		END IF %>
		<td><nobr><%= TRIM(RS1("projectName"))%></nobr></td>
		<td align="center"><INPUT size="5" maxlength="6" name="<%= RS1("userID")%>:<%= RS1("projectName")%>" value="<%= FormatNumber(RS1("phase1"),2)%>"></td>
		<td align="center"><INPUT size="5" maxlength="6" name="<%= RS1("userID")%>:<%= RS1("projectName")%>" value="<%= FormatNumber(RS1("phase2"),2)%>"></td>
		<td align="center"><INPUT size="5" maxlength="6" name="<%= RS1("userID")%>:<%= RS1("projectName")%>" value="<%= FormatNumber(RS1("phase3"),2)%>"></td>
		<td align="center"><INPUT size="5" maxlength="6" name="<%= RS1("userID")%>:<%= RS1("projectName")%>" value="<%= FormatNumber(RS1("phase4"),2)%>"></td>
		<td align="center"><INPUT size="5" maxlength="6" name="<%= RS1("userID")%>:<%= RS1("projectName")%>" value="<%= FormatNumber(RS1("phase5"),2)%>"></td></tr>
<%		RS1.MoveNext
	LOOP %>
	<tr><td colspan=7 align="center"><br><br><input type="submit"><br><br></td></tr>
	</FORM>
</table>
</body>
</html><%
RS1.Close
Set connUsers = Nothing
connSWPPP.Close
Set connSWPPP = Nothing %>