<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("../loginUser.asp")
End If
%> 
<!-- #include file="../../connSWPPP.asp" --> 

<%
If Request.Form.Count > 0 Then
	for n = 0 to 999 step 1
		if Request("chk:"& CStr(n)) = "on" then
			id = Request("id:"& CStr(n))
			rights = Request("rights:"& CStr(n))
			Response.Write("Updating: ID:" & id & " = " & rights & "</br>")
			'update the database
			supdate = "UPDATE Users SET rights='"& rights & "' WHERE userID="& id 
			'Response.Write(supdate)
			SET update = connSWPPP.Execute(supdate)
		end if
	next
End If
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Cleanup Users for Admins</title>
	<link rel="stylesheet" href="../../../global.css" type="text/css">
	<script type="text/javascript">
  function check_all_items(obj) {
    for (i=0; i<999; i++){
        var name = "chk:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = 'on';
            s[0].checked = true;
        } else {
            break;
        }
    }
  }

  function uncheck_all_items(obj){
     for (i=0; i<999; i++){
        var name = "chk:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = 'off';
            s[0].checked = false;
        } else {
            break;
        }
     }
  }
  </script>
</head>
<body>
<!-- #INCLUDE FILE="../../adminHeader3.inc" -->  
<form action="<%= Request.ServerVariables("script_name") %>" method="post">
<input type="submit" value="Update Users">
<input type="button" value="check all Items" onclick="check_all_items(this)" />
<input type="button" value="un-check all Items" onclick="uncheck_all_items(this)" />
<table><tr>
<th>Update</th>
<th>UserID</th>
<th>First Name</th>
<th>Last Name</th>
<th>Active</th>
<th>User</th>
<th>Action</th>
<th>Email</th>
<th>Erosion</th>
<th>VSCR</th>
<th>LDSCR</th>
<th>Inspector</th>
<th>Director</th>
<th>Admin</th>
<th>DB Rights</th>
<th>Rights</th>
<% 
SQLSELECT = "SELECT * FROM Users ORDER BY lastName"
'Response.Write(SQLSELECT & "<br>")
Set connUsers = connSWPPP.Execute(SQLSELECT)
recCount = 0
updateCount = 0

DO WHILE Not connUsers.EOF
	recCount = recCount + 1
	firstName = Trim(connUsers("firstName"))
	lastName = Trim(connUsers("lastName"))
	userID = Trim(connUsers("userID"))
	rights = Trim(connUsers("rights"))
	active = connUsers("active")

	SQL1 = "SELECT p.*, u.userID, u.firstName, u.lastName, u.rights as rights1, pu.rights as rights2" &_
		" FROM Projects as p LEFT JOIN ProjectsUsers as pu ON p.projectID=pu.projectID LEFT JOIN Users as u" &_
		" ON pu.userID=u.userID WHERE p.phaseNum=1 AND u.userID="& userID &" ORDER BY projectName ASC, projectPhase ASC"
		'Response.Write(SQL1)
	SET RS1=connSWPPP.execute(SQL1)
	
	'---	Initialize loop variables -------------------------------------------------------------------
	compCount=0 
	userChecked=""
	insChecked=""
	dirChecked=""
	actChecked=""
	eroChecked=""
	vscrChecked=""
	ldscrChecked=""
	recEmailChecked=""
	adminChecked=""

	If rights = "admin" THEN
		adminChecked="YES"
	End If

	DO WHILE NOT RS1.EOF 
		'---	For each Project User Record, by project ID, compare userID to this user --------------------
		'---	On a match, find out the rights checkbox that needs to be set -------------------------------
		IF TRIM(RS1("userID"))=userID THEN
			'--Response.Write(RS1("userID") &":"& userID &":"& RS1("rights2") &":"& RS1("emailReport") &"<br>")
			SELECT CASE TRIM(RS1("rights2"))
				CASE "user"			userChecked="YES"
				CASE "inspector"	insChecked="YES"
				CASE "email"		recEmailChecked="YES"
				CASE "ecc"		   recEmailChecked="YES"
				CASE "bcc"		   recEmailChecked="YES"
				CASE "action"		actChecked="YES"
				CASE "erosion"		eroChecked="YES"
				CASE "vscr"       vscrChecked="YES"
				CASE "ldscr"      ldscrChecked="YES"
				CASE "director"	
					dirName= RS1("firstName") &" "& RS1("lastName") 
					dirChecked="YES"
				CASE "admin"      adminChecked="YES"
			END SELECT
		END IF
		'---	If this record does not match the current user, check to find out if it is ------------------
		'---	the director for this Project. If it is save director values for it -------------------------
		IF TRIM(RS1("rights2"))="director" THEN 
			dirName= TRIM(RS1("firstName")) &" "& TRIM(RS1("lastName")) 
		END IF
		RS1.MoveNext
	LOOP
	if adminChecked = "YES" then
		current_rights = "admin"
	elseif dirChecked = "YES" then
		current_rights = "director"
	elseif insChecked = "YES" then
		current_rights = "inspector"
	elseif ldscrChecked = "YES" then
		current_rights = "ldscr"
	elseif vscrChecked = "YES" then
		current_rights = "vscr"
	elseif eroChecked = "YES" then
		current_rights = "erosion"
	elseif recEmailChecked = "YES" then
		current_rights = "email"
	elseif actChecked = "YES" then
		current_rights = "action"
	elseif userChecked = "YES" then
		current_rights = "user"
	else
		current_rights = "user"
	end if
	
	if current_rights <> rights then
		updateCount = updateCount + 1
		%>
		<tr>
		<td><input type="checkbox" name="chk:<%=updateCount-1%>" /></td>
		<td><input type="text" name="id:<%=updateCount-1%>" value="<%=userID%>" /></td>
		<td><%=firstName%></td>
		<td><%=lastName%></td>
		<td><%=active%></td>
		<td><%=userChecked%></td>
		<td><%=actChecked%></td>
		<td><%=recEmailChecked%></td>
		<td><%=eroChecked%></td>
		<td><%=vscrChecked%></td>
		<td><%=ldscrChecked%></td>
		<td><%=insChecked%></td>
		<td><%=dirChecked%></td>
		<td><%=adminChecked%></td>
		<td><%=rights%></td>
		<td><input type="text" name="rights:<%=updateCount-1%>" value="<%=current_rights%>" /></td>
		</tr>
		<%
	else
		update = ""
	end if

	connUsers.MoveNext
LOOP %>
</table>
</form>
</body>
</html>