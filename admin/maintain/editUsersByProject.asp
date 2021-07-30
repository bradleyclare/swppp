<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

Server.ScriptTimeout=4500

projectID = Request("pID")

%> <!-- #include file="../connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	highestRights="user"
	
   'delete all user rights for this project
	SQLDELETE = "DELETE FROM ProjectsUsers WHERE projectID=" & projectID
	'Response.Write(SQLDELETE & "<br/>")
   connSWPPP.execute(SQLDELETE)

   '-----	rightsValue="000000000" '-- user,action,erosion,email,CC,BCC,inspector,director,admin -----------------
	rightsValue="000000000"
   ' ----------------------- Inspector, Director, User, Action, Email in Projects User  -------- 
	prevUserID = 0
   For Each Item in Request.Form
      slen = Len(item)
      userID = Right(item,slen-4)
		Select Case Left(Item,3)
			Case "use"
				rights="user"
			Case "act"
				rights="action"
			Case "ero"
				rights="erosion"
			Case "emr"
				rights="email"
			Case "ecc"
				rights="ecc"
			Case "bcc"
				rights="bcc"
			Case "vsc"
				rights="vscr"
			Case "lds"
				rights="ldscr"
			Case "ins"
				rights="inspector"
			Case "dir"
				rights="director"
		End Select
		If rights<>"" then
         SQL3 = "INSERT INTO ProjectsUsers (projectID, userID, rights) VALUES (" & projectID & ", " & userID &", '"& rights & "')"
         'Response.Write(SQL3 & "<br/>")
			connSWPPP.Execute(SQL3)
		end if 'item=inspector, director or user
      rights=""
	Next
End If

SQL1="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2 & "<br/>")
Set RS1=connSWPPP.execute(SQL1) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Edit Users by Project</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<body>
   <!-- #include file="../adminHeader2.inc" -->
	<form action="editUsersByProject.asp?pID=<%=projectID%>" method="post" >
      <h1>User Rights for <%= RS1("projectName") %>&nbsp;<%= RS1("projectPhase")%></h1>
      <!--- ----------------------------------------- Rights --------------------------------------- --->
      <center><input type="submit" value="Update Rights"></center>
      <br /><br />
      <table width="100%" border="0">
		<tr><th>id</th><th>name</th><th>user</th><th>email to</th><th>cc</th><th>bcc</th>
      <% If (Session("validDirector") or Session("validAdmin")) then '- directors can create action managers %>	
			<th>action</th>	
			<th>erosion</th>
			<th>vscr</th>
			<th>ldscr</th>
      <% End If
      If Session("validAdmin") then %>
			<th>director</th>
			<th>inspector</th>	
      <% End If 'Session("validAdmin") %>
      </tr>
      <% SQLSELECT = "SELECT userID, firstName, lastName FROM Users WHERE active=1 ORDER BY firstName, lastName"
      'response.Write(SQLSELECT)
      Set rsUser = connSWPPP.Execute(SQLSELECT)
      
      Do While Not rsUser.EOF
         userID = rsUser("userID")
         firstName = Trim(rsUser("firstName"))
         lastName = Trim(rsUser("lastName")) 
         
         SQL2="SELECT rights FROM ProjectsUsers WHERE userID="& userID & " AND projectID =" & projectID
         'response.Write(SQL2)
         Set RS2=connSWPPP.execute(SQL2)

         userChecked=False
         insChecked=False
         recEmailChecked=False
         recCCChecked=False
         recBCCChecked=False
         actChecked=False
         eroChecked=False
         dirChecked=False
			vscrChecked=False
			ldscrChecked=False
         Do While Not RS2.EOF
            SELECT CASE TRIM(RS2("rights"))
			      CASE "user"			userChecked=True
			      CASE "inspector"	insChecked=True
			      CASE "email"		recEmailChecked=True
			      CASE "ecc"		   recCCChecked=True
			      CASE "bcc"		   recBCCChecked=True
			      CASE "action"		actChecked=True
			      CASE "erosion"		eroChecked=True
			      CASE "vscr"       vscrChecked=True
					CASE "ldscr"      ldscrChecked=True
					CASE "director"	
				      dirName= firstName &" "& lastName 
				      dirChecked=True
		      END SELECT
            RS2.MoveNext
         Loop 
         RS2.Close %>     

         <tr><td><%=userID %></td>
            <td><%=firstName %>&nbsp<%=lastName %></td>
            <td><input type="checkbox" name="use:<%= userID %>"
			   <% If userChecked then %>checked class="checked" <% End If %>
			   >u</td>
            <td><input type="checkbox" name="emr:<%= userID %>"
			   <% If recEmailChecked then %>checked class="checked" <% End If %>
			   >em</td>
            <td><input type="checkbox" name="ecc:<%= userID %>"
			   <% If recCCChecked then %>checked class="checked"  <% End If %>
			   >c</td>
            <td><input type="checkbox" name="bcc:<%= userID %>"
			   <% If recBCCChecked then %>checked class="checked"  <% End If %>
			   >b</td>
         <% If (Session("validDirector") or Session("validAdmin")) then '- directors can create action managers %>	
			   <td><input type="checkbox" name="act:<%= userID %>" disabled="disabled" 
			   <% If actChecked then %>checked class="checked"  <% End If %>
			   >a</td>	
			   <td><input type="checkbox" name="ero:<%= userID %>"
			   <% If eroChecked then %>checked class="checked"  <% End If %>
			   >er</td>
				<td><input type="checkbox" name="vsc:<%= userID %>"
			   <% If vscrChecked then %>checked class="checked"  <% End If %>
			   >vscr</td>
				<td><input type="checkbox" name="lds:<%= userID %>"
			   <% If ldscrChecked then %>checked class="checked"  <% End If %>
			   >ldscr</td>
         <%	End If
         If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
			   <td><input type="checkbox" name="dir:<%= userID %>"
			   <% If dirChecked then %>checked class="checked"  <% End If %>
			   >d</td>
			   <td><input type="checkbox" name="ins:<%= userID %>"
			   <% If insChecked then %>checked class="checked"  <% End If %>
			   >i</td>	
         <% End If 'Session("validAdmin") %>
         </tr>
         <% rsUser.MoveNext
      LOOP %>
      </table>
		<br><br>
      <center><input type="submit" value="Update Rights"></center>
	</form>
</body>
</html>
<%'--	Release Resources ---------------------------------------------------------------------------
rsUser.Close
Set rsUser = Nothing
RS1.Close
Set RS1 = Nothing %>