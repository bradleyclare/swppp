<%@ Language="VBScript" %>

<%
currentDate = date()
endDate = currentDate
startDate=DateAdd("d",-30,currentDate)
%><!-- #include file="../connSWPPP.asp" --><%
SQL0="SELECT * FROM Projects WHERE active=1 ORDER BY projectName ASC"
SET RS0=connSWPPP.execute(SQL0)

%>
<html>
<head>
<title>SWPPP INSPECTIONS : Check Inspection Reports</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<%	

text = ""
late_cnt = 0
If RS0.EOF Then
	text = text & "<h2>Sorry no projects found.</h2>"
Else
	text = "<h1>late reports</h1><table>"
	text = text & "<tr><th>inspec date</th><th>project name</th><th>report type</th><th>inspector</th><th>days late</th></tr>"
	Do While Not RS0.EOF
		projID = RS0("projectID")
		projectName = Trim(RS0("projectName")) & "&nbsp;" & Trim(RS0("projectPhase"))
		SQL1="SELECT TOP 1 * FROM Inspections WHERE projectID="& projID &" ORDER BY inspecDate DESC"
		SET RS1=connSWPPP.execute(SQL1)

		If not RS1.EOF Then
			inspecID = RS1("inspecID")
			reportSQLSELECT = "SELECT inspecID, inspecDate, reportType" & _
				", firstName, lastName, i.projectID, i.projectName, i.projectPhase, released" & _
				" FROM Inspections as i, Users as u" & _
				" WHERE inspecID = " & inspecID & " AND i.userID = u.userID"
			'Response.Write(reportSQLSELECT & "<br>")
			Set rsReports = connSWPPP.execute(reportSQLSELECT)
			
			inspecDate = rsReports("inspecDate")
			reportType = Trim(rsReports("reportType"))
			userFullName = Trim(rsReports("firstName")) & "&nbsp;" & Trim(rsReports("lastName"))
			
			date_thresh = 9999
			if StrComp(reportType,"Weekly") = 0 Then
				date_thresh = 7
			elseif StrComp(reportType,"Bi-Weekly") = 0 Then
				date_thresh = 14
			elseif StrComp(reportType,"Rainfall") = 0 Then
				date_thresh = 14
			elseif StrComp(reportType,"Storm Event") = 0 Then
				date_thresh = 14
			elseif StrComp(reportType,"Site Visit") = 0 Then
				date_thresh = 7
			elseif StrComp(reportType,"Monthly") = 0 Then
				date_thresh = 30
			End if

			date_diff = datediff("d",inspecDate,currentDate) 

			'response.write(inspecDate & " : " & projectName & " : " & reportType & " : " & date_thresh & " : " & date_diff & "<br/>")

			If date_diff > date_thresh Then
				late_cnt = late_cnt + 1
				text = text & "<tr><td>" & inspecDate & "</td>" & _
					"<td>" & projectName & "</td>" & _
					"<td>" & reportType & "</td>" & _
					"<td>" & userFullName & "</td>" & _ 
					"<td>" & date_diff - date_thresh & "</td>"
			End If
		End If

		RS0.MoveNext
	Loop
	text = text & "</table>"

	if late_cnt = 0 Then
		text = text & "<h3>No late reports found over the past year</h3>"
	End If

End If ' END No Results Found
	
Response.Write(text)

rsReports.Close
Set rsReports = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
</body>
</html>