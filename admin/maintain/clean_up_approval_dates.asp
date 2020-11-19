<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If 
%>
<!-- #include file="../connSWPPP.asp" -->
<html>
<head>
	<title>SWPPP INSPECTIONS : Clean Up Approval Dates</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Clean Up Approval Dates</h1>
<% approvalSQLSELECT = "SELECT * FROM HortonApprovals"
'Response.Write(approvalSQLSELECT)
Set rsApproval = connSWPPP.execute(approvalSQLSELECT)
prev_date = ""
Do While Not rsApproval.EOF 
	id = rsApproval("id")
	app_date = TRIM(rsApproval("date")) 
	'Response.Write(id & " - " & app_date & "</br>")
	If app_date = "1/1/1900" Then
		approvalSQLUPDATE = "UPDATE HortonApprovals SET date = '" & prev_date & "' WHERE id = " & id
		Response.Write(approvalSQLUPDATE & "</br>")
		connSWPPP.execute(approvalSQLUPDATE)
   Else
		prev_date = app_date
	End If
	rsApproval.MoveNext
Loop 
%>
</body>
</html>