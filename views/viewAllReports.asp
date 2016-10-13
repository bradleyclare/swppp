<%@ Language="VBScript" %>
<%
'--	Determine the month to report on -----------------------------------------------------------
startDate=Request("startDate")
IF (TRIM(startDate)="" OR NOT(IsDate(startDate)) )Then 
	startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
Else
	startDate=CDATE(Month(CDATE(Request("startDate")))&"/1/"&Year(CDATE(Request("startDate"))))
End IF
endDate=DateAdd("m",1,startDate)
endDate=DateAdd("d",-1,endDate)
startMonth=Month(startDate)
startYear=Year(startDate)
recordOrd = Request("orderBy")
If recordOrd = "" Then recordOrd = " p.projectName ASC, inspecDate DESC" End If
%><!-- #include file="../admin/connSWPPP.asp" --><%
reportSQLSELECT = "SELECT inspecID, inspecDate, projectCounty" & _
	", i.projectID, p.projectName" & _
	" FROM Inspections as i, Projects as p" & _
	" WHERE i.projectID = p.projectID" &_
	" AND i.inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
	" AND i.projectID IN ( SELECT projectID FROM ProjectsUsers" 
IF NOT Session("validAdmin") THEN
	reportSQLSELECT=reportSQLSELECT & " WHERE  ProjectsUsers.userID = " & Session("userID") 
END IF
	reportSQLSELECT=reportSQLSELECT & ") ORDER BY " & recordOrd
Set rsReports = connSWPPP.execute(reportSQLSELECT) %>
<html>
<head>
<title>SWPPP INSPECTIONS : View Inspection Reports</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body>
<!-- #include file="../header2.inc" -->
<table width="90%" border="0">
	<tr valign="middle"> 
		<TD align=left colspan=5><h1>View Inspection Reports</h1></TD></tr>
<% If Session("validAdmin") then %>
	<tr valign=top><TD align=left colspan=3 valign="middle">
			<FORM action="qbXLS.asp" method="post">
				Create Spreadsheet for <%= MonthName(startMonth)%> with starting invoice number 
				<input type="text" name="iNum" value="" maxlength="6" size="4">
				<input type="hidden" name="xDate" value="<%= startDate%>">
				<input type="submit" value="go">&nbsp;
				<button style="border-width:thin; font-size:xx-small; height:17px; width:17px; background-color:#e5e6e8;" 
					onClick="alert('IMPORT HELP\n\n1. Enter the next invoice number from Quickbooks.When you click the \'go\' button\nthe system will create a spreadsheet that you can import into Quickbooks.\n\n2. \'Save As\' the generated spreadsheet as a tab delimited file. Use any file-name\nyou need, but enclose the file-name in quotation marks and end the file-name\nwith \'.iff\' before saving. This will allow the file to be imported into Quickbooks.\n   (for example: \'\'testfile.iff\'\')\n\n3. In Quickbooks, import the file by accessing \'File>>Utilities>>Import IIF file\'.\nThis will import the invoices into the system.');">?</button></FORM></TD>
		<TD align=right colspan=2>
<% ELSE %><TD align=right colspan=5><% END IF %>View Reports for month of:&nbsp;
			<SELECT name="startMonth" onChange="navigateMe(this.value);"><%	
	m=DateDiff("m",startDate,Date)+6
'	IF Month(startDate)=Month(Date) THEN m=-13 ELSE m=-8 END IF
	FOR n= -m to (m-6) step 1 
		optDate=DateAdd("m",n,startDate)
		optMonth=MonthName(Month(optDate))
		optYear=Year(optDate) %>
			<OPTION value="<%= optDate%>"<% IF Month(optDate)=startMonth THEN%> selected<% END IF%>><%= optMonth%>, <%= optYear%>
<%	Next %>
			</SELECT>
		</TD>
	</tr>
	<tr>
<!--		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=projectCounty&startDate=<%=startDate%>"><b>County</b></a></th>-->
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=p.projectName, inspecDate DESC&startDate=<%=startDate%>"><b>Project</b></a></th>
		<th><a href="<%= Request.ServerVariables("script_name") %>?orderBy=inspecDate DESC&startDate=<%=startDate%>"><b>Date</b></a></th>
	</tr>
	<tr> 
		<td colspan="3"><img src="../../images/dot.gif" width="5" height="5"></td>
	</tr>
<%
	If rsReports.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>Sorry " & _
			"no current data available for the Month requested.</i></b></td></tr>")
	Else
		altColors = "#e5e6e8"
		
		Do While Not rsReports.EOF
			inspecDate = rsReports("inspecDate")
			projectName = Trim(rsReports("projectName"))
			projectCounty = Trim(rsReports("projectCounty"))
%>
	<tr align="center" bgcolor="<% = altColors %>"> 
<!--		<td nowrap><% = projectCounty %></td>-->
		<td nowrap><% = projectName %></td>
		<td><% = inspecDate %></td>
	</tr>
<%
			' Alternate Row Colors
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			
			rsReports.MoveNext
		Loop
		
	End If ' END No Results Found
	
rsReports.Close
Set rsReports = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
	<tr> 
		<td colspan="3">&nbsp;</td>
	</tr>
</table>
</body>
<script language="VBScript">
function navigateMe(param)
	window.navigate("viewAllReports.asp?startDate="&param)
end function
</SCRIPT>
</html>
