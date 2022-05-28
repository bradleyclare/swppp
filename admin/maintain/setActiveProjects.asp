<%Response.Buffer = False%>
<%
If Not Session("validAdmin") AND not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If
%> <!-- #include file="../connSWPPP.asp" --> <%

SQL0="SELECT * FROM Projects ORDER BY projectName ASC"
SET RS0=connSWPPP.execute(SQL0) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Calculate Active Projects</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<h1>Calculate Active Projects</h1>
<hr />
    <% cnt=0
	currentDate = Date()
	Response.Write("<table><tr><th>Project Name</th><th>Last Inspec Date</th><th>Days Since Last Report</th><th>Status</th></tr>")
	DO WHILE NOT RS0.EOF 
	    'find most recent inspect report for this projectName
		projID = RS0("projectID")
		projName = RS0("projectName")
		projPhase = RS0("projectPhase")
		activeFlag = RS0("active")
		IF activeFlag=true THEN
			SQL1="SELECT TOP 1 * FROM Inspections WHERE projectID="& projID &" ORDER BY inspecDate DESC"
			SET RS1=connSWPPP.execute(SQL1)
			IF RS1.EOF THEN
				SQL2="UPDATE Projects SET active=0 WHERE projectID="& projID 
				SET RS2=connSWPPP.execute(SQL2)
				Response.Write("<tr><td>" & projName & " " & projPhase & "</td><td>No Report</td><td>n/a</td><td>Set Inactive</td></tr>")
			END IF
			DO WHILE NOT RS1.EOF
				diff = DateDiff("d",RS1("inspecDate"),currentDate)
				IF diff > 90 THEN
					SQL2="UPDATE Projects SET active=0 WHERE projectID="& projID 
					SET RS2=connSWPPP.execute(SQL2)
					Response.Write("<tr><td>" & projName & " " & projPhase & "</td><td>" & RS1("inspecDate") & "</td><td>" & diff & "</td><td>Set Inactive</td></tr>")
				ELSE
					Response.Write("<tr><td>" & projName & " " & projPhase & "</td><td>" & RS1("inspecDate") & "</td><td>" & diff & "</td><td>Active</td></tr>")
				END IF
				RS1.moveNext
			LOOP
			RS1.Close
			Set RS1 = Nothing 
		Else
			'uncomment these lines to reset the active database
			'SQL2="UPDATE Projects SET active=1 WHERE projectID="& projID 
			'SET RS2=connSWPPP.execute(SQL2)
		END IF
		RS0.moveNext    
	LOOP 
	Response.Write("</table>")
	RS0.Close
	Set RS0 = Nothing
	connSWPPP.Close
	Set connSWPPP = Nothing %>
	</form>
</table>
</body>
</html>