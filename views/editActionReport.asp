<%@ Language="VBScript" %>
<!-- #include file="../admin/connSWPPP.asp" --><% 
If 	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")	
End If
IF Request("pID")="" OR Request("ID")="" THEN Response.Write("<SCRIPT language=VBScript>window.close()</SCRIPT>") END IF
projectID = Request("pID")
actionID = Request("ID")
DEL = Request("DEL")
SQL0="SELECT * FROM ProjectsUsers WHERE userID="& Session("userID") &" AND rights='action' AND projectID="& projectID 
SET RS0=connSWPPP.execute(SQL0)
IF RS0.eof THEN
	IF NOT(Session("validAdmin") OR Session("validDirector")) THEN
		RS0.Close
		Set RS0=nothing
		'connSWPPP.Close
		'SET conSWPPP=nothing 
		Response.Write("<SCRIPT language=VBScript>window.close()</SCRIPT>")
	END IF
END IF
someError=False
If DEL = "1" Then
    Response.Redirect("deleteActionReport.asp?pID="& projectID &"&ID="& actionID)
End IF 
If Request.Form.Count > 0 Then
	IF 	(NOT(Session("userID")="")) AND _
			(IsNumeric(Session("userID"))) AND _
			(IsNumeric(Request("pID"))) AND _
			(IsNumeric(actionID)) AND _
		(NOT(TRIM(Request("actionText"))="")) THEN
			actionText=CleanText(Request("actionText"))
			SQL1="UPDATE Actions SET ActionText = LEFT('"& actionText &"',1000), last_actionDate='"& Date() &"', last_userID="& Session("userID") &_
				" WHERE actionID="& actionID &" AND ActionText<>LEFT('"& actionText &"',1000)"
			connSWPPP.execute(SQL1)
			connSWPPP.Close
			SET connSWPPP=nothing
			Response.redirect("actionReport.asp?pID="& projectID)
	END IF
	someError=True
END IF 
canDelete=False
IF Session("validAdmin") THEN canDelete=True END IF
SQL2="SELECT * FROM ProjectsUsers pu inner join Actions a on pu.projectID=a.projectID" &_
	" WHERE pu.userID="& Session("userID") &" AND a.actionID="& actionID
SET RS2=connSWPPP.execute(SQL2)
IF NOT(RS2.EOF) THEN	
	DO WHILE NOT RS2.EOF
		IF TRIM(RS2("rights"))="director" AND Session("userID")=RS2("userID") THEN canDelete=True END IF
		IF RS2("orig_userID")=Session("userID") AND (TRIM(RS2("rights"))="action" or TRIM(RS2("rights"))="erosion") THEN canDelete=True END IF
		RS2.MoveNext
	LOOP
END IF
%><!-- #include file="cleaner.vb" --><%
SQL1="SELECT actionID, orig_actionDate, last_actionDate, dbo.fnGetFullName(orig_userID) as fullName1," &_
	" dbo.fnGetFullName(last_userID) as fullName2, LTRIM(RTRIM(actionText)) as actionText " &_
	" FROM Actions a WHERE a.actionID="& Request("ID") %>
<!-- #include file="../admin/connSWPPP.asp" --><%
SET RS1=connSWPPP.execute(SQL1) %>
<html><head>
<title>SWPPP INSPECTIONS - Add Actions Taken Report Entry</title>
<link rel="stylesheet" type="text/css" href="../global.css"></head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
<% IF someError THEN %><p><FONT size="+1" color="red">There was an error in either the date field or the TextBox.</FONT></p><% END IF %>
<font size="+1"><b>Action Report Entry</b></font><hr noshade size="1" width="90%">
<% IF canDelete THEN %>
	<button style="height:18px; width:120px;" onClick="window.location.href='deleteActionReport.asp?pID=<%=projectID%>&ID=<%=actionID%>&DEL=1'">
		<font size="-2">Delete Action Report</font></button><br><br>
<% END IF %>
<FORM action="<% = Request.ServerVariables("script_name") %>" method="post">
<INPUT type="hidden" name="pID" value="<%=projectID%>">
<INPUT type="hidden" name="ID" value="<%=actionID%>">
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="90" align=left>Creator</th>
		<th width="80" align=left>Date Created</th>
		<th width="130" align=left>Last Modified</th>
		<th align=left>Edit Action Taken</th></tr>
	<tr><td valign="top" align=left><%= Trim(RS1("fullName1")) %></TD>
		<td valign="top" align=left><%= Trim(RS1("orig_actionDate")) %></TD>
		<td valign="top" align=left><%= Trim(RS1("last_actionDate")) %><br>by<br><%= Trim(RS1("fullName2")) %></TD>
		<td align="left"><TEXTAREA cols="60" rows="10" name="actionText"><%= Trim(UnCleanText(RS1("actionText"))) %></TEXTAREA></TD></tr>
</table><br><br>
<input type="submit" Value="Edit Action Report"></FORM></center>
</body></html>