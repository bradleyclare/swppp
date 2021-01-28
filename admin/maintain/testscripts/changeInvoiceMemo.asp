<% Response.buffer = false
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../loginUser.asp")
End If 

%> <!-- #INCLUDE FILE="../../connSWPPP.asp" --> <%
'	Response.Write(SQLSELECT & "<br>")
SQL0="SELECT * FROM Projects"
SET RS0=connSWPPP.Execute(SQL0) 

If RS0.EOF Then
	noMatch = True
	Response.Write("No projects found.<br/>")
Else
	Response.Write("Processing Projects.<br/>")
	Do While Not RS0.EOF
		recCount = recCount + 1
		if RS0("invoiceMemo") = "Thank you for your business." Then
			SQL1="UPDATE Projects SET invoiceMemo='Due upon receipt.' WHERE projectID="& RS0("projectID") 
			SET RS1=connSWPPP.execute(SQL1)
			Response.Write(recCount & " - " & RS0("projectName") & " : " & RS0("invoiceMemo") & " CHANGED.<br/>")
		End If
		RS0.MoveNext
	Loop
	connSWPPP.Close
	Set connSWPPP = Nothing
End If ' no projects found
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>SWPPP :: Change Invoice Memo</title>
	<link rel="stylesheet" href="../../../global.css" type="text/css">
</head>
<body> 
<!-- #INCLUDE FILE="../../adminHeader3.inc" -->  
Done.
</body>
</html>
