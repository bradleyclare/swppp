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
	DO WHILE NOT RS0.EOF 
	    'find most recent inspect report for this projectName
		IF RS0("phaseNum")=1 THEN
			projID = RS0("projectID")
			SQL1="SELECT TOP 1 * FROM Inspections WHERE projectID="& projID &" ORDER BY inspecDate DESC"
			SET RS1=connSWPPP.execute(SQL1)
			IF RS1.EOF THEN
				SQL2="UPDATE Projects SET phaseNum=0 WHERE projectID="& projID 
				SET RS2=connSWPPP.execute(SQL2)
				Response.Write(RS0("projectName") & " " & RS0("projectPhase") & " : No Report --- UPDATED <br/>")
			END IF
			DO WHILE NOT RS1.EOF
				diff = DateDiff("m",RS1("inspecDate"),currentDate)
				IF diff > 3 THEN
					SQL2="UPDATE Projects SET phaseNum=0 WHERE projectID="& projID 
					SET RS2=connSWPPP.execute(SQL2)
					Response.Write(RS0("projectName") & " " & RS0("projectPhase") & " : " & RS1("inspecDate") & " : " & diff & " --- UPDATED <br/>")
				ELSE
					Response.Write(RS0("projectName") & " " & RS0("projectPhase") & " : " & RS1("inspecDate") & " : " & diff & "<br/>")
				END IF
				RS1.moveNext
			LOOP 
		END IF
		RS0.moveNext    
	LOOP 
	Response.Write("DONE<br><br>")
	RS0.Close
	Set RS0 = Nothing
	RS1.Close
	Set RS1 = Nothing
	connSWPPP.Close
	Set connSWPPP = Nothing %>
	</form>
</table>
</body>
</html>