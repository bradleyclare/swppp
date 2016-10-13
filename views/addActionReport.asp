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
IF Request("pID")="" THEN self.close() END IF
projectID = Request("pID")
SQL0="SELECT * FROM ProjectsUsers WHERE "& Session("userID") &" IN (SELECT userID FROM ProjectsUsers WHERE rights in ('action','erosion') AND projectID="& projectID &")"
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
success=False
If Request.Form.Count > 0 Then
	IF 	(NOT(Session("userID")="")) AND _
			(IsNumeric(Session("userID"))) AND _
		(NOT(Request("pID")="")) AND _
			(IsNumeric(Request("pID"))) AND _
			(IsDate(Request("actionDate"))) AND _
		(NOT(TRIM(Request("actionText"))="")) THEN
			SQL1="INSERT INTO Actions (projectID, orig_userID, last_userID, orig_actionDate, last_actionDate, ActionText) VALUES ("&_
				projectID &", "& Session("userID") &", "& Session("userID") &", '"& Request("actionDate") &"', '"& Request("actionDate") &"', LEFT('"& CleanText(Request("actionText")) &"',1000))"
			connSWPPP.execute(SQL1)
			connSWPPP.Close
			SET connSWPPP=nothing
			success=True
		'Response.Write("<SCRIPT language=VBScript>window.close()</SCRIPT>")
	ELSE
	someError=True
	END IF
END IF

function CleanText(textStr)
	CleanText=REPLACE(textStr,"/*","")
	CleanText=REPLACE(CleanText,"*/","")
	CleanText=REPLACE(CleanText,chr(34),"&quot;")
	CleanText=REPLACE(CleanText,chr(39),"&apos;")
	CleanText=REPLACE(CleanText,chr(45),"&hyphen;")
	CleanText=REPLACE(CleanText,chr(145),"&lsquo;")
	CleanText=REPLACE(CleanText,chr(147),"&ldquo;")
	CleanText=REPLACE(CleanText,chr(169),"&copy;")
	CleanText=REPLACE(CleanText,chr(174),"&reg;")
end function %>
<html><head>
<title>SWPPP INSPECTIONS - Add Actions Taken Report Entry</title>
<link rel="stylesheet" type="text/css" href="../global.css"></head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
<% IF someError THEN %><p><FONT size="+1" color="red">There was an error in either the date field or the TextBox.</FONT></p><% END IF %>
<% IF success THEN %><p><FONT size="+1" color="red">Action recorded successfully.</FONT></p><% END IF %>
<font size="+1"><b>Action Report Entry</b></font><hr noshade size="1" width="90%">
<FORM action="<% = Request.ServerVariables("script_name") %>" method="post">
<INPUT type="hidden" name="pID" value="<%=projectID%>">
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="100" align=left>Date</th><th align=left>Action Taken</th></tr>
	<tr><td valign="top" align=left><INPUT type="text" name="actionDate" value="<%= Date()%>"></TD>
		<td align="left"><TEXTAREA cols="60" rows="5" name="actionText"></TEXTAREA></TD></tr>
</table><br><br>
<input type="submit" Value="Submit Action Report"></FORM></center>
</body></html>