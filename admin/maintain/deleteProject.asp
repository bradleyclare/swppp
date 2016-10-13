<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") &_
		"?" & Request("query_string")
	Response.Redirect("loginUser.asp")
End If
projID = Request("id")
IF (IsNull(projID) OR NOT(IsNumeric(projID))) THEN Response.Redirect("viewProjects.asp") END IF
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	SQLDELETE = "DELETE FROM Projects WHERE projectID=" & projID
	' Response.Write(coordSQLINSERT & "<br><br>")
	connSWPPP.Execute(SQLDELETE)	
	connSWPPP.Close
	Set connSWPPP = Nothing
	Response.Redirect("viewProjects.asp")
End If 
SQL0="SELECT i.inspecDate, i.projectName, i.projectPhase, u.lastName, u.firstName FROM Inspections i, Users u" &_
	" WHERE i.userID=u.userID AND i.projectID="& projID &" ORDER BY i.inspecDate DESC"
SET RS0=connSWPPP.Execute(SQL0) %>
<html><head>
	<title>SWPPP INSPECTIONS : Delete Project</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<script language="JavaScript" src="../js/validCoordinates.js"></script>
	<script language="JavaScript" src="../js/validCoordinates1.2.js"></script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Delete Project</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0"><%
IF RS0.BOF AND RS0.EOF THEN %>
<form action="<%= Request.ServerVariables("script_name") %>" method="post">
	<input type="hidden" name="id" value="<%= projID %>">
	<tr><td align="center"><font color="#FF0000">You are about to permanently Delete this Project.<br>
	You can return to <a href="../default.asp">Admin Home</a> to abort this operation.<br>
	Do you want to continue this delete process?</font>	
	<br><br><input type="submit" value="Yes, Delete"></td></tr>
</form><%
ELSE %>
	<TR><TD><font size="+1" color="red">This project still has the following Inspections associated with it.<br>It can not be deleted until these inspections are either<br>deleted or associated with other projects.<br><br></font></td></tr> 
<%	DO WHILE NOT RS0.EOF %>
	<tr><td>on <b><%= RS0("inspecDate")%></b> for <%= RS0("projectName")%> <%= RS0("projectPhase")%> by <%= RS0("firstName")%> <%= RS0("lastName")%></td></tr><%
		RS0.MoveNext
	LOOP
END IF %>	
</table></body></html><%	
connSWPPP.Close
Set connSWPPP = Nothing %>