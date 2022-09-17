<%@ Language="VBScript" %>
<% Response.Buffer = False
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

userID = TRIM(Request("IDuser")) '--- name userID would conflict with user rights names
listUserID = TRIM(Session("userID")) '--- the user ID used to determine what projects are visible.
companyID = Request("companyID")
%> <!-- #include file="../connSWPPP.asp" --> <%
processed=false
If Request.Form.Count > 0 Then
	highestRights="user"
	SQLU="SELECT rights FROM Users WHERE userID="& userID
	SET RSU= connSWPPP.execute(SQLU)
	highestRights=TRIM(RSU("rights"))
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	Function titleCase(strValue)
		strValue=Ucase(Left(strValue,1)) & Lcase(Mid(strValue,2,Len(strValue)-1))
		titleCase=Replace(strValue,"'","''")
	end function
	trimmedQualifications=REPLACE(Request("qualifications"),"'","#@#")
    if Request("seeScoring") = "on" then seeScoring = 1 Else seeScoring = 0 End If
    if Request("openItemAlerts") = "on" then openItemAlerts = 1 Else openItemAlerts = 0 End If
    if Request("repeatItemAlerts") = "on" then repeatItemAlerts = 1 Else repeatItemAlerts = 0 End If
    userGroupString = Request("userGroup")
    userGroupArray = Split(userGroupString,"-")
	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors
		SQLUPDATE =	"UPDATE Users SET" & _
			" firstName = '" & titleCase(Request("firstName")) & "'" & _
			", lastName = '" & titleCase(Request("lastName")) & "'" & _
			", email = '" & Request("email") & "'" & _
			", phone = '" & Request("phone") & "'" & _
			", pswrd = '" & Request("pswrd") & "'" & _
			", signature = '" & Request("signature") & "'" & _
			", noImages = '" & Request("noImages") & "'" & _
			", seeScoring = '" & seeScoring & "'" & _
			", openItemAlerts = '" & openItemAlerts & "'" & _
			", repeatItemAlerts = '" & repeatItemAlerts & "'" & _
			", qualifications = '" & trimmedQualifications & "'" 
	' -----------------------------------------------------------------
			If Len(userGroupString) > 0 Then
			SQLUPDATE = SQLUPDATE & ", userGroupID = " & CInt(Trim(userGroupArray(0)))
		End If
	' --------------------------- Admin User --------------------------
			If Request("admin")="on" then 
				SQLUPDATE = SQLUPDATE & ", rights= 'admin'"
				highestRights="admin"
			ELSE
				SQLUPDATE = SQLUPDATE & ", rights= 'user'"
			End If
		' -----------------------------------------------------------------
		SQLUPDATE = SQLUPDATE & " WHERE userID = " & userID
		'Response.Write(SQLUPDATE)
		connSWPPP.Execute(SQLUPDATE)
	Else
		SQLUPDATE =	"UPDATE Users SET" & _
			" firstName = '" & titleCase(Request("firstName")) & "'" & _
			", lastName = '" & titleCase(Request("lastName")) & "'" & _
			", phone = '" & Request("phone") & "'"
		' -----------------------------------------------------------------
		SQLUPDATE = SQLUPDATE & " WHERE userID = " & userID
		'Response.Write(SQLUPDATE)
		connSWPPP.Execute(SQLUPDATE)
	End If

	SQLDELETE = "DELETE FROM ProjectsUsers WHERE userID=" & userID
	IF Session("validDirector") THEN 
		SQLDELETE=SQLDELETE & " AND projectID IN (SELECT projectID FROM ProjectsUsers" &_
				" WHERE userID=" & listUserID &")"
	END IF
	connSWPPP.execute(SQLDELETE)

'-----	rightsValue="00000000000" '-- user,action,erosion,email,CC,BCC,vscr,ldscr,inspector,director,admin -----------------
	If Request("admin")="on" then rightsValue= "00000000001" else rightsValue="00000000000" End If
'response.write(rights & "<br/>")
' ----------------------- Inspector, Director, User, Email in Projects User  -------- 
		For Each Item in Request.Form
            If Item <> "userGroup" Then
			    Select Case Left(Item,3) 'loop through and replace the appropriate character with 1 for certain rights, mid function returns substring
				    Case "use"
					    rights="user"
					    rightsValue= "1"& MID(rightsValue,2)
				    Case "act"
					    rights="action"
					    rightsValue=MID(rightsValue,1,1) &"1"& MID(rightsValue,3)
				    Case "emr"
					    rights="email"
					    rightsValue=MID(rightsValue,1,2) &"1"& MID(rightsValue,4)
				    Case "ecc"
					    rights="ecc"
					    rightsValue= MID(rightsValue,1,3) &"1"& MID(rightsValue,5)
				    Case "bcc"
					    rights="bcc"
					    rightsValue= MID(rightsValue,1,4) &"1"& MID(rightsValue,6)
				    Case "ero"
					    rights="erosion"
					    rightsValue= MID(rightsValue,1,5) &"1"& MID(rightsValue,7)
					 Case "vsc"
					    rights="vscr"
					    rightsValue=MID(rightsValue,1,6) &"1"& MID(rightsValue,8)
					 Case "lds"
					    rights="ldscr"
					    rightsValue=MID(rightsValue,1,7) &"1"& MID(rightsValue,9)
					 Case "ins"
					    rights="inspector"
					    rightsValue=MID(rightsValue,1,8) &"1"& MID(rightsValue,10)
					 
    '					SQLa="SELECT projectName FROM Projects WHERE projectID="& Request(Item)
    '					SET RSa=connSWPPP.execute(SQLa)
    '					IF NOT(RSa.BOF AND RSa.EOF) THEN 
    '					    projName=RSa(0) 
    '					ELSE 
    '					    projName="error" 
    '					END IF
					    SQLP="SELECT * FROM Commissions WHERE userID="& userID &" AND projectID ='"& Request(Item) &"'"
					    SET RSP=connSWPPP.execute(SQLP)
					    IF RSP.EOF THEN
						    phase1=20
						    phase2=10
						    phase3=5
						    phase4=0
						    phase5=30
					    ELSE
						    phase1=RSP("phase1")
						    phase2=RSP("phase2")
						    phase3=RSP("phase3")
						    phase4=RSP("phase4")
						    phase5=RSP("phase5")
					    END IF
					    RSP.Close
					    SET RSP=nothing
    '					RSa.close
    '					SET RSa=nothing
					    SQL0=SQL0 &" EXEC sp_UpdateCommissions "& userID &", '"& projName &"', "& phase1 &", "& phase2 &", "& phase3 &", "& phase4 &", "& phase5
				    Case "dir"
					    rights="director"
					    rightsValue=MID(rightsValue,1,9) &"1"& MID(rightsValue,11)
			    End Select
			    If rights<>"" then
                    exeCmd = "sp_InsertPU "& userID &", "& Request(Item) &", '"& rights &"'"
                    'Response.Write(exeCmd & "<br/>")
				    connSWPPP.Execute(exeCmd)
			    end if 'item=inspector, director or user
			    rights=""
            End If
		Next
		'response.write("Rights Value:" & rightsValue & "<br/>")
		highestRights="user"
		FOR n = 1 to 11 step 1
			IF (MID(rightsValue,n,1)="1") THEN 
				SELECT CASE n
					CASE 1	highestRights="user"
					CASE 2	highestRights="action"
					CASE 3	highestRights="email"
					CASE 4   highestRights="email"
					CASE 5   highestRights="email"
					CASE 6   highestRights="erosion"
					CASE 7   highestRights="vscr"
					CASE 8   highestRights="ldscr"
					CASE 9	highestRights="inspector"
					CASE 10	highestRights="director"
					CASE 11	highestRights="admin"
					CASE ELSE highestRights=highestRights
				END SELECT
			END IF
		NEXT 
		'response.write("Highest Rights:" & highestRights & "</br>")
		IF highestRights <>"" THEN
			connSWPPP.execute("UPDATE Users SET rights='"& highestRights &"' WHERE userID="& userID)
		End If
'--		Clean UP Commissions Table -------------------------------------------------------
'		SQL0="DELETE FROM Commissions WHERE commID NOT IN(" &_
'			" SELECT c.commID FROM Commissions c JOIN Projects p ON" &_
'			" c.projectName=p.projectName JOIN ProjectsUsers pu ON" &_ 
'			" p.projectID=pu.projectID AND c.userID=pu.userID" &_
'			" WHERE pu.rights='inspector' and pu.userID="& userID &")" &_
'			" AND userID="& userID
'response.Write(SQL0)
'		connSWPPP.execute(SQL0)	
		processed="true"
End If
	SQLSELECT = "SELECT * FROM Users WHERE userID = " & userID
	Set rsUser = connSWPPP.Execute(SQLSELECT)
	
	dateEntered = rsUser("dateEntered")
	firstName = Trim(rsUser("firstName"))
	lastName = Trim(rsUser("lastName"))
	email = Trim(rsUser("email"))
	phone = Trim(rsUser("phone"))
	pswrd = Trim(rsUser("pswrd"))
	signature = Trim(rsUser("signature"))
	noImages = TRIM(rsUser("noImages"))
	qualifications= TRIM(rsUser("qualifications"))
    seeScoring = rsUser("seeScoring")
    openItemAlerts = rsUser("openItemAlerts")
    repeatItemAlerts = rsUser("repeatItemAlerts")
    userGroupID = rsUser("userGroupID")
	IF IsNull(qualifications) THEN qualifications="" END IF
	rsUser.Close
	Set rsUser = Nothing

'If processed THEN Response.Redirect("viewUsersAdmin.asp") END IF %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Edit User</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
	<script language="JavaScript" src="../js/validUsers.js"></script>
   <script language="JavaScript" src="../js/validUsers1.2.js"></script>
   <script type="text/javascript">
      function check_all_bcc(obj) {
        for (i=0; i<999; i++){
            var name = "coord:complete:" + i.toString();
            var s = document.getElementsByName(name);
            if (s.length > 0){
                s[0].value = 'on';
                s[0].checked = true;
            } else {
                break;
            }
        }
      }

      function uncheck_all_bcc(obj){
         for (i=0; i<999; i++){
            var name = "coord:complete:" + i.toString();
            var s = document.getElementsByName(name);
            if (s.length > 0){
                s[0].value = 'off';
                s[0].checked = false;
            } else {
                break;
            }
         }
      }

      function check_all_inspector(obj) {
         for (i=0; i<999; i++){
            var name = "coord:complete:" + i.toString();
            var s = document.getElementsByName(name);
            if (s.length > 0){
               s[0].value = 'on';
               s[0].checked = true;
            } else {
               break;
            }
         }
      }

      function uncheck_all_inspector(obj){
         for (i=0; i<999; i++){
            var name = "coord:complete:" + i.toString();
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
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<form action="<%= Request.ServerVariables("script_name") %>" method="post" onSubmit="return isReady(this)";>
	<input type="hidden" name="IDuser" value="<%= userID %>">
		<tr><td colspan="2"><h1>Edit User</h1></td></tr>
        <tr><td colspan="2" align="center"><input type="button" value="Delete User" 
			onClick="location='deleteUser.asp?IDuser=<%= Request("IDuser") %>'; return false";></td></tr>
		<tr><td width="35%" align="right">date entered:</td>
			<td width="65%"><%= dateEntered %></td></tr>
		<tr><td width="35%" align="right">first name:</td>
			<td width="65%"><input type="text" name="firstName" size="20" maxlength="20" 
				value="<%= firstName %>"></td></tr>
		<tr><td align="right">last name:</td>
			<td><input type="text" name="lastName" size="20" maxlength="20" 
				value="<%= lastName %>"></td></tr>
		<tr><td align="right">phone:</td>
			<td><input type="text" name="phone" size="15" maxlength="15" value="<%= phone %>"></td></tr>
<% If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>		
		<tr><td align="right">email:</td>
			<td><input type="text" name="email" size="30" maxlength="50" value="<%= email %>"></td></tr>
		<tr><td align="right">password:</td>
			<td><input type="password" name="pswrd" size="15" maxlength="15"
				value="<%= pswrd %>">&nbsp;&nbsp;<input type="button" value="View" 
					onClick="alert('Password: ' + form.pswrd.value)";></td></tr>
		<tr><td align="right">signature file:</td>
			<td><select name="signature">

<%' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") &"images\signatures\")
Set gifDirectory = objFolder.Files

For Each gifFile In gifDirectory
	fileName = gifFile.Name %>
					<option value="<% = fileName %>"<% If signature = fileName Then %> selected<% End If %>><% = fileName %></option>
<% Next
Set objFolder = Nothing
Set gifDirectory = Nothing %>
				</select>&nbsp;&nbsp;<input type="button" value="upload signature file" 
					onClick="location='upSigEditUser.asp?userID=<%= userID %>'; return false";></td></tr>
		<tr><td align="right">view images:</td>
			<td><input type="radio" name="noImages" value="0"<% IF noImages=0 THEN %> checked<% END IF%>>Yes
				<input type="radio" name="noImages" value="1"<% IF noImages=1 THEN %> checked<% END IF%>>No</td></tr>
		<tr><td align="right">see scoring:</td>
            <td><input type="checkbox" name="seeScoring" 
            <% If (seeScoring) = True Then %>
                checked
            <% End if %>
            /></td>
		</tr>
        <tr><td align="right">receive open item alerts:</td>
            <td><input type="checkbox" name="openItemAlerts" 
            <% If (openItemAlerts) = True Then %>
                checked
            <% End if %>
            /></td>
		</tr>
        <tr><td align="right">receive repeat item alerts:</td>
            <td><input type="checkbox" name="repeatItemAlerts" 
            <% If (repeatItemAlerts) = True Then %>
                checked
            <% End if %>
            /></td>
		</tr>
        <tr><td align="right" valign=top>qualifications:</td>
			<td><TEXTAREA cols="50" rows="3" name="qualifications"><%= REPLACE(qualifications,"#@#","'") %></TEXTAREA></td></tr>
		<tr><td align="right">User Group:</td>
        <td><select name="userGroup"
           <% If not Session("validAdmin") then %>
           readonly 
           <% End If %>
           ><option value="0 - No Group">0 - No Group</option>
        <% SQLSELECT = "SELECT userGroupID, userGroupName FROM UserGroups"
        'Response.Write(SQLSELECT & "<br>")
        Set connGroups = connSWPPP.Execute(SQLSELECT)
        Do While Not connGroups.EOF %>
            <option value="<%=connGroups("userGroupID")%> - <%=connGroups("userGroupName")%>" 
            <% If userGroupID = connGroups("userGroupID") Then %> 
                selected 
            <% End If %>
            ><%=connGroups("userGroupID")%> - <%=connGroups("userGroupName")%></option>
            <% connGroups.MoveNext 
        LOOP %>
     </select></td></tr>
<% END IF %>
    <tr><td><input type="submit" value="Update User"></td></tr>
</table>
<br /><br />
<!---<table><tr>
   <td><input type="button" value="Check all BCC" onclick="check_all_bcc(this)" /></td>
   <td><input type="button" value="UnCheck all BCC" onclick="uncheck_all_bcc(this)" /></td>
   <td><input type="button" value="Check all Inspector" onclick="check_all_inspector(this)" /></td>
   <td><input type="button" value="UnCheck all Inspector" onclick="uncheck_all_inspector(this)" /></td>
</tr></table>
<br /><br />-->
<!--- ----------------------------------------- Rights --------------------------------------- --->
<table width="100%" border="0">
		<tr><td align="left"><br><font size="+1">Rights</font><br><br></td></tr>		
<% 	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
<% 	SQLSELECT = "SELECT rights FROM Users" & _
		" WHERE userID = "& userID &" AND rights='admin'"
	Set adminRights = connSWPPP.Execute(SQLSELECT) %>
		<tr><td align="right">Admin:</td>
			<td align=center><input type="checkbox" name="admin" 
				<% If not adminRights.eof then %>Checked<% end if %>><br></td></tr><% 	
	adminRights.close
	Set adminRights = Nothing
	end If %>		
		<tr><th>project name</th>
			<th>user</th>
			<th>email to</th>
			<th>cc</th>
<% 	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
            <th>bcc</th>
<%  end if 'Session("validAdmin">) %>			
<% 	If (Session("validDirector") OR Session("validAdmin")) then '- directors can create action managers %>		
			<th>erosion</>
			<th>VSCR</th>
			<th>LDSCR</th>
<%	End If
 	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
			<th>director</th>
			<th>inspector</th></tr>	
<% end if 'Session("validAdmin")
SQL1 = "SELECT p.*, u.userID, u.firstName, u.lastName, u.rights as rights1, pu.rights as rights2" &_
	" FROM Projects as p LEFT JOIN ProjectsUsers as pu ON p.projectID=pu.projectID LEFT JOIN Users as u" &_
	" ON pu.userID=u.userID"
	IF Session("validDirector") AND NOT(Session("validAdmin")) THEN 
		SQL1=SQL1 & " WHERE p.active=1 AND p.projectID IN (SELECT projectID FROM ProjectsUsers" &_
		" WHERE userID=" & listUserID &" AND rights='director')"
	ELSE
		SQL1=SQL1 & " WHERE p.active=1"
	END IF
SQL1=SQL1 & " ORDER BY projectName ASC, projectPhase ASC"
'Response.Write(SQL1)
SET RS1=connSWPPP.execute(SQL1)
'---	Initialize loop variables -------------------------------------------------------------------
compCount=0 
dispProjID=RS1("projectID")
dispProjName=TRIM(RS1("projectName"))&" "& TRIM(RS1("projectPhase"))
userChecked=False
insChecked=False
dirChecked=False
actChecked=False
eroChecked=False
vscrChecked=False
ldscrChecked=False
recEmailChecked=False
recCCChecked=False
recBCCChecked=False
dirName="None"
altColors="#ffffff"
'---	Begin the Loop ------------------------------------------------------------------------------
DO WHILE NOT RS1.EOF 
'---	For each Project User Record, by project ID, compare userID to this user --------------------
'---	On a match, find out the rights checkbox that needs to be set -------------------------------
	IF TRIM(RS1("userID"))=userID THEN
'--Response.Write(RS1("userID") &":"& userID &":"& RS1("rights2") &":"& RS1("emailReport") &"<br>")
		SELECT CASE TRIM(RS1("rights2"))
			CASE "user"			userChecked=True
			CASE "email" 		recEmailChecked=True 
			CASE "ecc"		    recCCChecked=True
			CASE "bcc"		    recBCCChecked=True
			CASE "inspector"	insChecked=True
			CASE "action"		actChecked=True
			CASE "erosion"		eroChecked=True
			CASE "vscr"       vscrChecked=True
			CASE "ldscr"      ldscrChecked=True
			CASE "director"	
				dirName= RS1("firstName") &" "& RS1("lastName") 
				dirChecked=True
		END SELECT
	END IF
'---	If this record does not match the current user, check to find out if it is ------------------
'---	the director for this Project. If it is save director values for it -------------------------
	IF TRIM(RS1("rights2"))="director" THEN dirName= TRIM(RS1("firstName")) &" "& TRIM(RS1("lastName")) END IF
	RS1.MoveNext
	IF NOT RS1.EOF THEN	currProjID=RS1("projectID") END IF
	IF (dispProjID <> currProjID) OR (RS1.EOF) THEN
'---	All records for this Project have been checked. We now have to display the available --------
'---	Checkboxes and then move on to the next Project ---------------------------------------------
		compCount=compCount+1 %>
	<tr bgcolor="<%= altColors %>"><td><%= dispProjName %></td>
<!--	check box for User --------------------------------------------------------------------- --->
		<td align=center><input type="checkbox" name="use<%= compCount %>" value="<%= dispProjID %>"
			<% If userChecked then %>checked <% End If %>
			onClick="if (!(document.forms[0].use<%=compCount%>.checked) && !(document.forms[0].admin.checked)) { (document.forms[0].emr<%=compCount%>.checked=false); }"></td>
		<td align=center><input type="checkbox" name="emr<%= compCount %>" value="<%= dispProjID %>"
			<% If recEmailChecked then %>checked <% END If%>
			onClick="if (!(document.forms[0].use<%=compCount%>.checked) && !(document.forms[0].admin.checked)) { (document.forms[0].emr<%=compCount%>.checked=false); }"></td>
		<td align=center><input type="checkbox" name="ecc<%= compCount %>" value="<%= dispProjID %>"
			<% If recCCChecked then %>checked <% END If%>
			onClick="if (!(document.forms[0].use<%=compCount%>.checked) && !(document.forms[0].admin.checked)) { (document.forms[0].emr<%=compCount%>.checked=false); }"></td>
<%	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
		<td align=center><input type="checkbox" name="bcc<%= compCount %>" value="<%= dispProjID %>"
			<% If recBCCChecked then %>checked <% END If%>
			onClick="if (!(document.forms[0].use<%=compCount%>.checked) && !(document.forms[0].admin.checked)) { (document.forms[0].emr<%=compCount%>.checked=false); }"></td>
<%  End If
 	If (Session("validDirector") OR Session("validAdmin")) then '- directors can create action managers %>		
<!--- ----------------------------------------- Erosion ----------------------------------------- --->
		<td align=center><input type="checkbox" name="ero<%= compCount %>" value="<%= dispProjID %>"
			<% If eroChecked then %>checked<% End If %>></td>
<!--- ----------------------------------------- VSCR -------------------------------------- --->
		<td align=center><input type="checkbox" name="vsc<%= compCount %>" value="<%= dispProjID %>"
			<% If vscrChecked then %>checked<% End If %>></td>
<!--- ----------------------------------------- LDVSCR -------------------------------------- --->
		<td align=center><input type="checkbox" name="lds<%= compCount %>" value="<%= dispProjID %>"
			<% If ldscrChecked then %>checked<% End If %>></td>
<% 	End If
	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>		
<!--- ----------------------------------------- Director --------------------------------------- --->
		<td align=center><input type="checkbox" name="dir<%= compCount %>" value="<%= dispProjID %>"
			<% If dirChecked then %>checked<% End If %>
			<% IF NOT(dirChecked) AND NOT(dirName="None") THEN%>readonly<%END IF%>></td>
<!--- ----------------------------------------- Inspector -------------------------------------- --->
		<td align=center><input type="checkbox" name="ins<%= compCount %>" value="<%= dispProjID %>"
			<% If insChecked then %>checked<% End If %>></td></tr>	
<% end if 'Session("validAdmin") 
'---- 	Reset the loop Variables --------------------------------------------------------------------
		IF NOT RS1.EOF THEN
			dispProjID=RS1("projectID")
			dispProjName=RS1("projectName")&" "& TRIM(RS1("projectPhase"))
			userChecked=False
			insChecked=False
			actChecked=False
			eroChecked=False
			vscrChecked=False
			ldscrChecked=False
			recEmailChecked=False
			recCCChecked=False
			recBCCChecked=False
			dirChecked=False
			dirName="None"
			If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
		END IF
	END IF
LOOP %>
		<TR><TD align="center" colspan=5><br><br><input type="submit" value="Update User"></TD></TR>
	</form>
</table>
</body>
</html>
<%'--	Release Resources ---------------------------------------------------------------------------
RS1.Close
Set RS1 = Nothing %>