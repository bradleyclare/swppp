<%@ Language="VBScript" %><%
'IF Request("pID")="" THEN Response.Write("<SCRIPT language='VBScript'>window.close()</SCRIPT>") END IF
projectID = Request("pID")
%><!-- #include file="../admin/connSWPPP.asp" --><%
SQL1="SELECT actionID, orig_actionDate, dbo.fnGetFullName(orig_userID) as fullName, LTRIM(RTRIM(actionText)) as actionText, orig_userID" &_
	" FROM Actions a WHERE a.projectID="& projectID &" ORDER BY orig_actionDate DESC"
SET RS1=connSWPPP.execute(SQL1)
SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
SET RS2=connSWPPP.execute(SQL2)

function UnCleanText(textStr)
	UnCleanText=REPLACE(textStr,"&quot;",chr(34))
	UnCleanText=REPLACE(UnCleanText,"&apos;",chr(39))
	UnCleanText=REPLACE(UnCleanText,"&hyphen;",chr(45))
	UnCleanText=REPLACE(UnCleanText,"&lsquo;",chr(145))
	UnCleanText=REPLACE(UnCleanText,"&ldquo;",chr(147))
	UnCleanText=REPLACE(UnCleanText,"&copy;",chr(169))
	UnCleanText=REPLACE(UnCleanText,"&reg;",chr(174))
end function %>
<html>
<head>
<STYLE>
tr.highlighted {
	cursor:hand;
	background-color:silver
}
</STYLE>
<title>SWPPP INSPECTIONS - Actions Taken Report for <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></title>
<link rel="stylesheet" type="text/css" href="../global.css">
</head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
<font size="+1"><b>Actions Taken Report<br>for <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font><hr noshade size="1" width="90%">
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="100" align=left>Date</th><th align=left>Action Taken</th></tr>
<% 	DO WHILE NOT RS1.EOF 
		canEdit=false
		IF Session("validAdmin") OR Session("validDirector") OR Session("userID")=RS1(4) THEN canEdit=true END IF %>
	<tr <% IF canEdit THEN %>onMouseOver="this.className='highlighted';" onMouseOut="this.className='';" onClick="window.location='editActionReport.asp?ID=<%=RS1(0) %>&pID=<%= projectID %>';"<% END IF %>><td align=left><%= RS1(1)%><BR><%=Trim(RS1(2))%></TD><td align=left><%=Trim(UnCleanText(RS1(3)))%></TD></tr> 
<%		RS1.MoveNext
 	LOOP 
	RS1.Close
	SET RS1=nothing
	connSWPPP.Close
	SET connSWPPP=nothing %>
</center></table><br><br>
</body>
</html>