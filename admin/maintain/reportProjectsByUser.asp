<%Response.Buffer = False%>
<%
If Not Session("validAdmin") AND not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If
%> <!-- #include file="../connSWPPP.asp" --> <%

group = Request.QueryString("group")
If isempty(group) Then
    group = "A"
End If

'SQL0="SELECT * FROM Users ORDER BY lastName, firstName ASC"
'SET RS0=connSWPPP.execute(SQL0) 

if Session("validAdmin") Then
   'select the companies for which this user is a valid Director
   SQLSELECT = "SELECT projectID" & _
		   " FROM ProjectsUsers"
   Set connComp = connSWPPP.Execute(SQLSELECT)
Else
   'select the companies for which this user is a valid Director
   SQLSELECT = "SELECT projectID" & _
		   " FROM ProjectsUsers" &_
		   " WHERE userID=" & Session("userID") &_
		   " AND rights='director'"
   Set connComp = connSWPPP.Execute(SQLSELECT)
End If
' select users who have rights to those companies
SQLSELECT = "SELECT DISTINCT u.userID, firstName, lastName, pu.rights" &_
	" FROM Users as u JOIN ProjectsUsers as pu" &_
	" ON u.userID=pu.userID JOIN Projects as p" &_
	" ON pu.projectId=p.projectID" &_
	" WHERE u.userID = pu.userID AND pu.rights IN ('director', 'user')"  &_
	" AND u.rights!='admin' AND p.projectID IN (" 
Do while not connComp.eof
	if not subsequent then 'first time
		SQLSELECT = SQLSELECT & connComp("projectID")
		subsequent=true
	else
		SQLSELECT = SQLSELECT & ", " & connComp("projectID")
	end if
	connComp.movenext
Loop

SQLSELECT = SQLSELECT & ") ORDER BY lastName"

'Response.Write(SQLSELECT & "<br>")
Set RS0 = connSWPPP.Execute(SQLSELECT)

connComp.Close
Set connComp = Nothing
   
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Report Projects by User</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<h1>Report Projects by User</h1>
<hr/>
<h3>Select Letter to List Available Users</h3>
<ul class="list-inline">
    <% lastletter = ""
    DO WHILE NOT RS0.EOF 
        lastName = Trim(RS0("lastName"))
        letter = LCase(Left(lastName, 1))
        isnumber = isnumeric(letter) or StrComp(letter,"(") = 0
        If not isnumber and StrComp(letter,lastletter) <> 0 Then
            lastletter = letter %>
    <li><a href="reportProjectsByUser.asp?group=<% =letter %>"><% =UCase(letter) %></a></li>
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
	<td><img align=bottom src='..\..\images\E.gif' height="12"> - Erosion</td>
</table>
<br><br>
<table width="100%" border="0">
	<tr><th width="50%"><b>Project</b></th>
	<th><b>Project Name</b></th>
	<th><b>Project Phase</b></th>
	<th><b>Rights</b></th></TR>
	<form name=form1>
   <% cnt=0
   DO WHILE NOT RS0.EOF
      userID = Trim(RS0("userID"))
      letter = LCase(Left(RS0("lastName"), 1))
      numbergroup = StrComp(group,"0-9") = 0 '0 if equal
      isnumber = isnumeric(letter) or StrComp(letter,"(") = 0
      If LCase(group) = letter or (numbergroup and isnumber) then	
         cnt = cnt + 1 %>
	      <tr bgcolor="<%= altColors1 %>"><td colspan=4 align=left><B><%= RS0("lastName")%>,&nbsp<%= RS0("firstName")%></B>&nbsp;
         <BUTTON type='button' id=btnShow<%=cnt%> style="background-color: red; height: 10px; width:10px;"
			      onclick="tbody<%=cnt%>.style.display=''; btnHide<%=cnt%>.style.display=''; btnShow<%=cnt%>.style.display='none';"></BUTTON>
		   <BUTTON type='button' id=btnHide<%=cnt%> style="display:none; background-color: green; height: 10px; width:10px;"
			      onclick="tbody<%=cnt%>.style.display='none'; btnShow<%=cnt%>.style.display=''; btnHide<%=cnt%>.style.display='none';"></BUTTON>
         <% SQL1 = "SELECT p.*, u.userID, u.firstName, u.lastName, u.rights as rights1, pu.rights as rights2" &_
	            " FROM Projects as p LEFT JOIN ProjectsUsers as pu ON p.projectID=pu.projectID LEFT JOIN Users as u" &_
	            " ON pu.userID=u.userID WHERE u.userID=" & userID & " ORDER BY projectName ASC, projectPhase ASC"
         '--Response.Write(SQL1)
         SET RS1=connSWPPP.execute(SQL1) %>
         <TBODY id="tbody<%=cnt%>" style="display: none;">
         <% altColors2 = "#F8F8FF"
         dispProjectName=""
		   DO WHILE NOT RS1.EOF
            currProjectName = RS1("projectName")&" "&RS1("projectPhase")
		 	   IF currProjectName<>dispProjectName THEN 
		 		   dispProjectName = currProjectName %>
               </td></tr>
               <tr bgcolor="<%= altColors2 %>"><td width=50%>&nbsp;</td>
               <td><%= dispProjectName %></td>
               <td><% 
				   If altColors2= "#F8F8FF" Then altColors2 = "#DCDCDC" Else altColors2 = "#F8F8FF" End If
			   END IF
            IF TRIM(RS1("userID"))=userID THEN
			      SELECT CASE TRIM(RS1("rights2")) 
				      CASE "admin"		%>&nbsp;<img align="bottom" src="..\..\images\Admin.gif"><%	
				      CASE "inspector"	%>&nbsp;<img align="bottom" src="..\..\images\I.jpg"><%
				      CASE "director"		%>&nbsp;<img align="bottom" src="..\..\images\D.jpg"><%
				      CASE "user"			%>&nbsp;<img align="bottom" src="..\..\images\U.jpg"><%	
				      CASE "action"		%>&nbsp;<img align="bottom" src="..\..\images\AM.gif"><%	
				      CASE "erosion"		%>&nbsp;<img align="bottom" src="..\..\images\E.gif"><%	
				      CASE "email"		%>&nbsp;<img align="bottom" src="..\..\images\email2.jpg"><%
				      CASE "ecc"		    %>&nbsp;<img align="bottom" src="..\..\images\CC.gif"><%
				      CASE "bcc"		    %>&nbsp;<img align="bottom" src="..\..\images\BCC.gif"><%
				   END SELECT 
            END IF
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