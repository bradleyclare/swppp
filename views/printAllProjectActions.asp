<%@ Language="VBScript" %>
<!-- #include virtual="admin/connSWPPP.asp" --><%
IF NOT(IsNumeric(Session("userID"))) THEN Response.Redirect("../default.asp") END IF
IF Session("validUser") THEN highestRights="user" END IF
IF Session("validInspector") THEN highestRights="ins" END IF
IF Session("validDirector") THEN highestRights="dir" END IF
IF Session("validAdmin") THEN highestRights="admin" END IF
SQL0="sp_getProjectsPhases "& Session("userID") &", '"& highestRights &"'"
SET RS0=connSWPPP.execute(SQL0)
curPID=Trim(Request("cPID"))
IF NOT(IsNumeric(curPID)) THEN curPID=-1 ELSE curPID=CINT(curPID) END IF

function UnCleanText(textStr)
	UnCleanText=REPLACE(textStr,"&quot;",chr(34))
	UnCleanText=REPLACE(UnCleanText,"&apos;",chr(39))
	UnCleanText=REPLACE(UnCleanText,"&hyphen;",chr(45))
	UnCleanText=REPLACE(UnCleanText,"&lsquo;",chr(145))
	UnCleanText=REPLACE(UnCleanText,"&ldquo;",chr(147))
	UnCleanText=REPLACE(UnCleanText,"&copy;",chr(169))
	UnCleanText=REPLACE(UnCleanText,"&reg;",chr(174))
end function
%><html>
<head>
<title>SWPPP INSPECTIONS : Print All Projects Actions</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body onLoad="window.print();">
<table width="95%" border=0 cellspacing=0 cellpadding=0>
<form name="form1" method="post">
	<tr valign="middle"> 
		<TD align=left colspan=3><h1>All Projects Actions</h1></TD></tr>
</form>
</table><div><table width="95%" border=0><%
cPID=-1
SQL1="sp_getAllProjectsActions "& Session("userID") &", "& curPID &", '"& highestRights &"'"
set RS1 = server.createobject("adodb.recordset")
RS1.open SQL1, connSWPPP
IF RS1.EOF THEN
%>	<tr><td colspan=3 width=600>There are no Action Reports for the Project/Phase that you selected.</td></tr><%
ELSE
	bColor="transparent"  
	DO WHILE NOT RS1.EOF
		IF cPID<>RS1("projectID") THEN
			cPID=RS1("projectID") %>
	</table></div><div style="page-break-after:always;"><table width="95%" border=0>
	<tr><td width=90></td><td width="210"></td><td width="300"></td></tr>
	<tr><td colspan=3 height="3px"><hr></td></tr>				
	<tr><th colspan=2 width="300" style=""><%= TRIM(RS1("projectName")) %>&nbsp;<%= TRIM(RS1("projectPhase")) %></th>
		<th width="300"></th></tr>
	<tr><th width="90" align=right>date</th>
		<th colspan=2 align=left width="510" style="padding-left:10px;">action taken</th></tr><%
			bColor="transparent"
		END IF
		actDate=TRIM(RS1("orig_actionDate"))
		actItem=Trim(UnCleanText(RS1("actionText")))
				actMgrName=Trim(RS1("firstName")) &" "& Trim(RS1("lastName")) %>
	<tr bgcolor="<%= bColor%>"><td width="90" align=right valign="top"><%= actDate %><br><%= actMgrName %></td>
		<td colspan=2 align=left width="510"><%= actItem%></td></tr><%
		IF bColor="transparent" THEN bColor="silver" ELSE bColor="transparent" END IF
		RS1.movenext	   
	LOOP
END IF %>
	<tr><td colspan=3 height="3px"><hr></td></tr>
</table>
</body>
</html>