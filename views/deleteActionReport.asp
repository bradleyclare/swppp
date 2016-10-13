<%@ Language="VBScript" %>
<!-- #include virtual="admin/connSWPPP.asp" --><% 
If 	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	'response.Write(Session("validAdmin") &":"&Session("validDirector") &":"& Session("validInspector") &":"& Session("validUser") &":"& Session("validErosion"))
	Response.Redirect("../admin/maintain/loginUser.asp")	
End If
IF Request("pID")="" OR Request("ID")="" THEN Response.Write("<SCRIPT language=VBScript>window.close()</SCRIPT>") END IF
projectID = Request("pID")
actionID = Request("ID")
SQL0="SELECT * FROM ProjectsUsers WHERE userID="& Session("userID") &" AND rights='action' AND projectID="& projectID 
SET RS0=connSWPPP.execute(SQL0)
IF RS0.eof THEN
	IF NOT(Session("validAdmin") OR Session("validDirector")) THEN
		RS0.Close
		Set RS0=nothing
		connSWPPP.Close
		SET conSWPPP=nothing 
		'Response.Write("<SCRIPT language=VBScript>window.close()</SCRIPT>")
	END IF
END IF
someError=False
If Request.Form.Count > 0 Then
	SQL0="DELETE FROM Actions WHERE actionID="& actionID
	connSWPPP.execute(SQL0)
	Response.redirect("actionReport.asp?pID="& projectID)
End If 
%><!-- #include file="cleaner.vb" --><%
SQL1="SELECT actionID, orig_actionDate, last_actionDate, dbo.fnGetFullName(orig_userID) as fullName1," &_
	" dbo.fnGetFullName(last_userID) as fullName2, LTRIM(RTRIM(actionText)) as actionText " &_
	" FROM Actions a WHERE a.actionID="& Request("ID") %>
<!-- #include virtual="admin/connSWPPP.asp" --><%
SET RS1=connSWPPP.execute(SQL1) %>
<html><head>
<title>SWPPP INSPECTIONS - Delete Actions Taken Report Entry</title>
<link rel="stylesheet" type="text/css" href="../global.css"></head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
<font size="+1"><b>Delete Action Report Entry</b></font><hr noshade size="1" width="90%">
<FORM action="<% = Request.ServerVariables("script_name") %>" method="post">
<INPUT type="hidden" name="pID" value="<%=projectID%>">
<INPUT type="hidden" name="ID" value="<%=actionID%>">
<p><FONT size="+1" color="red">You are about to delete this Action Report permanently. It can not be recovered once you delete it.</font></p>
<input type="submit" Value="Delete this Action Report"></FORM>
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="100" align=left>Creator</th>
		<th width="50" align=left>Creation Date</th>
		<th width="150" align=left>Last Modified on</th>
		<th align=left>Edit Action Taken</th></tr>
<tr><td valign="top" align=left><%= Trim(RS1("fullName1")) %></td>
		<td valign="top" align=left><%= Trim(RS1("orig_actionDate")) %></td>
		<td valign="top" align=left><%= Trim(RS1("last_actionDate")) %> by <%= Trim(RS1("fullName2")) %></td>
		<td align="left"><TEXTAREA cols="60" rows="5" name="actionText"><%= Trim(UnCleanText(RS1("actionText"))) %></TEXTAREA></td></tr>
</table><br><br>
</center>
</body></html>