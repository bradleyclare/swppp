<%@ Language="VBScript" %>
<!-- #include virtual="admin/connSWPPP.asp" -->
<% IF NOT(IsNumeric(Session("userID"))) THEN Response.Redirect("../../default.asp") END IF
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
end function %>
<html>
<head>
	<title>SWPPP INSPECTIONS : Actions Taken Summary</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body>
	<!-- #include virtual="header.inc" -->
	
	<form name="form1" method="post">
		<h3>Actions Taken Summary</h3>
		<div class="nine columns alpha">
			Project View:
			<SELECT name="cPID" style="font-size: small;" onChange="redirectMe(this.value);">
				<OPTION value="-1"
				<% IF curPID=-1 THEN%> 
					selected
				<%END IF%>
				>All Projects</OPTION>
				<% DO WHILE NOT RS0.EOF %>
					<OPTION value="<%= RS0("projectID")%>"<%IF curPID=RS0("projectID") THEN%> selected<%END IF%>><%=RS0("projectName")%>&nbsp;<%= RS0("projectPhase")%></OPTION><% RS0.MoveNext
				LOOP %>
			</SELECT>
		</div>
		<div class="three columns omega">
			<div class="side-link">
				<a href="printAllProjectActions.asp?cPID=<%= curPID%>">Print</a>
			</div>
		</div>
	</form>
<table width="95%" border=0 cellpadding=1 cellspacing=1>
<% cPID=-1
SQL1="sp_getAllProjectsActions "& Session("userID") &", "& curPID &", '"& highestRights &"'"
set RS1 = server.createobject("adodb.recordset")
RS1.CursorLocation=3 'client side
RS1.CursorType = 3 'static recordset
RS1.PageSize = 5  
RS1.open SQL1, connSWPPP
IF RS1.EOF THEN%>	
	<tr><td colspan=3 width=600>There are no Action Reports for the Project/Phase that you selected.</td></tr>
<% ELSE
	page = request("pg")
	if page <= 0 or page = "" then
		RS1.AbsolutePage = 1	  
	else 
		IF (CINT(page)>RS1.PageCount) THEN RS1.AbsolutePage=1 ELSE RS1.AbsolutePage=CINT(page) END IF
	end if	 
	bColor="transparent"  
	for i = 1 to 5 step 1
		if RS1.eof then
			exit for
		else
			IF cPID<>RS1("projectID") THEN
				cPID=RS1("projectID") %>
	<tr><td width=90></td><td width="210"></td><td width="300"></td></tr>
	<tr><td colspan=3 height="2px"><hr></td></tr>				
	<tr><th colspan=2 width="300"><%= TRIM(RS1("projectName")) %>&nbsp;<%= TRIM(RS1("projectPhase")) %></th>
		<td width="300"></td></tr>
	<tr><th width="90" align=right>Date</th>
		<th colspan=2 align=left width="510" style="padding-left:10px;">Action Taken</th></tr><%
				bColor="transparent"
			END IF
				actDate=TRIM(RS1("orig_actionDate"))
				actItem=Trim(UnCleanText(RS1("actionText")))
				actMgrName=Trim(RS1("firstName")) &" "& Trim(RS1("lastName")) %>
	<tr bgcolor="<%= bColor%>"><td width="90" align=right valign="top"><%= actDate %><br><%= actMgrName %></td>
		<td colspan=2 align=left width="510"><%= actItem%></td></tr><%
				IF bColor="transparent" THEN bColor="silver" ELSE bColor="transparent" END IF
		end if
		page = RS1.AbsolutePage
		RS1.movenext	   
	next
END IF %>
	<tr><td colspan=3 height="2px"><hr></td></tr>
	<tr><td align=left><% IF page>1 THEN %>
			<span onMouseOver="this.style.cursor='hand';this.style.textDecoration='underline';" 
			onMouseOut="this.style.cursor='default';this.style.textDecoration='none';" 
			onclick="location='viewAllProjectActions.asp?pg=<%= page-1 %>&cPID=<%= curPID%>'">&lt;&lt; prev page</span><% END IF %>&nbsp;</td>
		<td></td>
		<td align=right>&nbsp;<% IF page<RS1.PageCount THEN%>
			<span onMouseOver="this.style.cursor='hand';this.style.textDecoration='underline';" 
			onMouseOut="this.style.cursor='default';this.style.textDecoration='none';" 
			onclick="location='viewAllProjectActions.asp?pg=<%= page+1 %>&cPID=<%= curPID%>'">next page &gt;&gt;</span><% END IF %></td></tr>
</table>
</body>
<script language="javascript">
 function redirectMe(cPID){
	window.location='viewAllProjectActions.asp?pg=<%= page-1 %>&cPID='+ cPID }
</script>
</html>