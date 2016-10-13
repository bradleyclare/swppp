<%@ Language="VBScript" %>
<!-- #include virtual="admin/connSWPPP.asp" --><%
If 	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If
IF Request("pID")="" THEN 
	self.close() 
END IF
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
<html>
<head>
	<title>SWPPP INSPECTIONS - Add Actions Taken Report Entry</title>
	<link rel="stylesheet" type="text/css" href="../global.css">
</head>
<body>
	<center>
	<br/><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
	<% IF someError THEN %>
		<p class="error">There was an error in either the date field or the TextBox.</p>
	<% END IF %>
	<% IF success THEN %>
		<p class="error">Action recorded successfully.</p>
	<% END IF %>
	<h3>Action Report Entry</h3>
	<form action="<% = Request.ServerVariables("script_name") %>" method="post">
		<input type="hidden" name="pID" value="<%=projectID%>">
		<div class="four columns alpha">
			<h5>Date</h5>
			<input type="text" name="actionDate" value="<%= Date()%>">
		</div>
		<div class="eight columns omega">
			<h5>Action Taken</h5>
			<textarea rows="5" name="actionText"></textarea>
		</div>
		<button type="submit" >Submit Action Report</button>
	</form>
	</center>
</body>
</html>