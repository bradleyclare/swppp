<%Response.Buffer = False%>
<%
If Not Session("validAdmin") AND not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If
%> <!-- #include file="../connSWPPP.asp" --> <%

group = Request.QueryString("group")
If isempty(group) Then
    group = "0-9"
End If

SQL0="SELECT * FROM Projects WHERE projectID IN (SELECT DISTINCT projectID FROM Inspections)"
	IF Session("validDirector") AND NOT(session("validAdmin")) THEN		
		SQL0=SQL0 & " AND projectID IN (SELECT projectID FROM ProjectsUsers WHERE userID=" & Session("userID") &" AND rights='director')"
	END IF
	SQL0=SQL0 & "ORDER BY projectName ASC"
SET RS0=connSWPPP.execute(SQL0) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Report Users by Project</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<h1>Report Users by Project</h1>
<hr />
<h3>Select Letter to List Available Projects</h3>
<ul class="list-inline">
    <li><a href="reportUsersByProject.asp?group=0-9">0-9</a></li>
    <% lastletter = ""
    DO WHILE NOT RS0.EOF 
        projectName = Trim(RS0("projectName"))
        letter = LCase(Left(projectName, 1))
        isnumber = isnumeric(letter) or StrComp(letter,"(") = 0
        If not isnumber and StrComp(letter,lastletter) <> 0 Then
            lastletter = letter %>
    <li><a href="reportUsersByProject.asp?group=<% =letter %>"><% =UCase(letter) %></a></li>
    <% End If		
    RS0.MoveNext
    LOOP 
    RS0.MoveFirst%>
</ul>
<div class="cleaner"></div>
<hr />
<h3><% =Ucase(group) %></h3>
<TABLE border="0">
	<tr><td><b>KEY:</B></td>
	<td><img align=bottom src='..\..\images\email2.jpg' height="12"> - Email</td>
	<td><img align=bottom src='..\..\images\CC.gif' height="12"> - CC</td>
    <% IF session("validAdmin") THEN %>	
    <td><img align=bottom src='..\..\images\BCC.gif' height="12"> - BCC</td>
	<td><img align=bottom src='..\..\images\I.gif' height="12"> - Inspector</td>
    <% END IF %>
	<td><img align=bottom src='..\..\images\D.jpg' height="12"> - Director</td>
	<td><img align=bottom src='..\..\images\U.jpg' height="12"> - User</td>
	<td><img align=bottom src='..\..\images\AM.gif' height="12"> - Action MGR</td>
	<td><img align=bottom src='..\..\images\E.gif' height="12"> - Erosion</td>
	<td><img align=bottom src='..\..\images\vscr.gif' height="12"> - VSCR</td>
	<td><img align=bottom src='..\..\images\ldscr.gif' height="12"> - LDSCR</td>
</table>
<br><br>
<table width="100%" border="0">
	<tr><th width="50%"><b>Project</b></th>
	<th><b>First Name</b></th>
	<th><b>Last Name</b></th>
	<th><b>Rights</b></th></TR>
	<form name=form1>
    <% cnt=0
	DO WHILE NOT RS0.EOF 
	    letter = LCase(Left(RS0("projectName"), 1))
        numbergroup = StrComp(group,"0-9") = 0 '0 if equal
        isnumber = isnumeric(letter) or StrComp(letter,"(") = 0
        If LCase(group) = letter or (numbergroup and isnumber) then	
            cnt = cnt + 1 %>
	        <tr bgcolor="<%= altColors1 %>"><td colspan=4 align=left><B><%= RS0("projectName")%>&nbsp<%= RS0("projectPhase")%></B>&nbsp;
		    <BUTTON type='button' id=btnShow<%=cnt%> style="background-color: red; height: 10px; width:10px;"
			    onclick="tbody<%=cnt%>.style.display=''; btnHide<%=cnt%>.style.display=''; btnShow<%=cnt%>.style.display='none';"></BUTTON>
		    <BUTTON type='button' id=btnHide<%=cnt%> style="display:none; background-color: green; height: 10px; width:10px;"
			    onclick="tbody<%=cnt%>.style.display='none'; btnShow<%=cnt%>.style.display=''; btnHide<%=cnt%>.style.display='none';"></BUTTON>
            <% SQL1="SELECT DISTINCT u.userID, lastName, firstName, pu.rights" &_
			    " FROM Users as u, ProjectsUsers as pu, Projects as p" &_
			    " WHERE pu.projectID="& RS0("projectID") &" AND pu.userID=u.userID " 
			IF Session("validDirector") AND NOT(session("validAdmin")) THEN		
		        SQL1 = SQL1 & " AND pu.rights IN('email','director','user','action','erosion','ecc','vscr','ldscr') and u.userID not in (Select userID From Users Where rights in ('admin','inspector'))"
			END IF
		    SQL1 = SQL1 & " ORDER BY firstName, lastName, pu.rights desc"
		    SET RS1=connSWPPP.execute(SQL1) %>
			<TBODY id="tbody<%=cnt%>" style="display: none;">
            <% altColors2 = "#F8F8FF"
		    currUserID=0
		    DO WHILE NOT RS1.EOF 
			    IF currUserID<>RS1("userID") THEN 
				    currUserID=RS1("userID") %>
		            </td></tr>
			        <tr bgcolor="<%= altColors2 %>"><td width=50%>&nbsp;</td>
			        <td><%= RS1("firstName") %></td><td><%= RS1("lastName") %></td>
			        <td><% 
				    If altColors2= "#F8F8FF" Then altColors2 = "#DCDCDC" Else altColors2 = "#F8F8FF" End If
			    END IF
			    SELECT CASE TRIM(RS1("rights")) 
				    CASE "admin"		%>&nbsp;<img align="bottom" src="..\..\images\Admin.gif"><%	
				    CASE "inspector"	%>&nbsp;<img align="bottom" src="..\..\images\I.jpg"><%
				    CASE "director"		%>&nbsp;<img align="bottom" src="..\..\images\D.jpg"><%
				    CASE "user"			%>&nbsp;<img align="bottom" src="..\..\images\U.jpg"><%	
				    CASE "action"		%>&nbsp;<img align="bottom" src="..\..\images\AM.gif"><%	
				    CASE "erosion"		%>&nbsp;<img align="bottom" src="..\..\images\E.gif"><%	
				    CASE "email"		%>&nbsp;<img align="bottom" src="..\..\images\email2.jpg"><%
				    CASE "ecc"		    %>&nbsp;<img align="bottom" src="..\..\images\CC.gif"><%
				    CASE "bcc"		    %>&nbsp;<img align="bottom" src="..\..\images\BCC.gif"><%
					 CASE "vscr"		    %>&nbsp;<img align="bottom" src="..\..\images\vscr.gif"><%
					 CASE "ldscr"		    %>&nbsp;<img align="bottom" src="..\..\images\ldscr.gif"><%
				END SELECT 
	        RS1.moveNext
	        LOOP %></TBODY></td></tr>
        <% End If 'Letter loop
        If altColors1 = "#C0C0C0" Then altColors1 = "#ffffff" Else altColors1 = "#C0C0C0" End If
    RS0.moveNext    
	LOOP 
	RS0.Close
	Set RS0 = Nothing
	connSWPPP.Close
	Set connSWPPP = Nothing %>
	</form>
</table>
</body>
</html>