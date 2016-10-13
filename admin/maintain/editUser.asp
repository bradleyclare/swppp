<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

userID = TRIM(Request("IDuser")) '--- name userID would conflict with user rights names
listUserID = TRIM(Session("userID")) '--- the user ID used to determine what projects are visible.
companyID = Request("companyID")
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
processed=false
If Request.Form.Count > 0 Then
	highestRights="user"
	SQL1="SELECT rights FROM Users WHERE userID="& userID
	SET RS1= connSWPPP.execute(SQL1)
	highestRights=TRIM(RS1("rights"))
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	Function titleCase(strValue)
		strValue=Ucase(Left(strValue,1)) & Lcase(Mid(strValue,2,Len(strValue)-1))
		titleCase=Replace(strValue,"'","''")
	end function
	trimmedQualifications=REPLACE(Request("qualifications"),"'","#@#")
	SQLUPDATE =	"UPDATE Users SET" & _
		" firstName = '" & titleCase(Request("firstName")) & "'" & _
		", lastName = '" & titleCase(Request("lastName")) & "'" & _
		", email = '" & Request("email") & "'" & _
		", pswrd = '" & Request("pswrd") & "'" & _
		", signature = '" & Request("signature") & "'" & _
		", noImages = '" & Request("noImages") & "'" & _
		", qualifications = '" & trimmedQualifications & "'" 
' --------------------------- Admin User --------------------------
		If Request("admin")="on" then 
			SQLUPDATE = SQLUPDATE & ", rights= 'admin'"
			highestRights="admin"
		ELSE
			SQLUPDATE = SQLUPDATE & ", rights= 'user'"
		End If
' -----------------------------------------------------------------
		SQLUPDATE = SQLUPDATE & " WHERE userID = " & userID
	connSWPPP.Execute(SQLUPDATE)
	
	SQLDELETE = "DELETE FROM ProjectsUsers WHERE userID=" & userID
	IF Session("validDirector") THEN 
		SQLDELETE=SQLDELETE & " AND projectID IN (SELECT projectID FROM ProjectsUsers" &_
		" WHERE userID=" & listUserID &")"
	END IF

	connSWPPP.execute(SQLDELETE)
'-----	rightsValue="000000000" '-- user,action,erosion,email,CC,BCC,inspector,director,admin -----------------
		If Request("admin")="on" then rightsValue= "000000001" else rightsValue="000000000" End If
' ----------------------- Inspector, Director, User, Action, Email in Projects User  -------- 
		For Each Item in Request.Form
			Select Case Left(Item,3)
				Case "use"
					rights="user"
					rightsValue= "1"& MID(rightsValue,2)
				Case "act"
					rights="action"
					rightsValue=MID(rightsValue,1,1) &"1"& MID(rightsValue,3)
				Case "ero"
					rights="erosion"
					rightsValue=MID(rightsValue,1,2) &"1"& MID(rightsValue,4)
				Case "emr"
					rights="email"
					rightsValue= MID(rightsValue,1,3) &"1"& MID(rightsValue,5)
				Case "ecc"
					rights="ecc"
					rightsValue= MID(rightsValue,1,4) &"1"& MID(rightsValue,6)
				Case "bcc"
					rights="bcc"
					rightsValue= MID(rightsValue,1,5) &"1"& MID(rightsValue,7)
				Case "ins"
					rights="inspector"
					rightsValue=MID(rightsValue,1,6) &"1"& MID(rightsValue,8)
'					SQLa="SELECT projectName FROM Projects WHERE projectID="& Request(Item)
'					SET RSa=connSWPPP.execute(SQLa)
'					IF NOT(RSa.BOF AND RSa.EOF) THEN 
'					    projName=RSa(0) 
'					ELSE 
'					    projName="error" 
'					END IF
					SQL1="SELECT * FROM Commissions WHERE userID="& userID &" AND projectID ='"& Request(Item) &"'"
					SET RS1=connSWPPP.execute(SQL1)
					IF RS1.EOF THEN
						phase1=20
						phase2=10
						phase3=5
						phase4=0
						phase5=30
					ELSE
						phase1=RS1("phase1")
						phase2=RS1("phase2")
						phase3=RS1("phase3")
						phase4=RS1("phase4")
						phase5=RS1("phase5")
					END IF
					RS1.Close
					SET RS1=nothing
'					RSa.close
'					SET RSa=nothing
					SQL0=SQL0 &" EXEC sp_UpdateCommissions "& userID &", '"& projName &"', "& phase1 &", "& phase2 &", "& phase3 &", "& phase4 &", "& phase5
				Case "dir"
					rights="director"
					rightsValue=MID(rightsValue,1,7) &"1"& MID(rightsValue,9)
			End Select
			If rights<>"" then
				connSWPPP.Execute("sp_InsertPU "& userID &", "& Request(Item) &", '"& rights &"'")
			end if 'item=inspector, director or user
			rights=""
		Next
		FOR n = 1 to 9 step 1
			IF (MID(rightsValue,n,1)="1") THEN 
				SELECT CASE n
					CASE 1	highestRights="user"
					CASE 2	highestRights="action"
					CASE 3	highestRights="erosion"
					CASE 7	highestRights="inspector"
					CASE 8	highestRights="director"
					CASE 9	highestRights="admin"
					CASE ELSE highestRights=highestRights
				END SELECT
			END IF
		NEXT 
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
	SQLSELECT = "SELECT firstName, lastName, email" & _
		", pswrd, dateEntered, signature, noImages, qualifications" & _
		" FROM Users WHERE userID = " & userID
	Set rsUser = connSWPPP.Execute(SQLSELECT)
	
	dateEntered = rsUser("dateEntered")
	firstName = Trim(rsUser("firstName"))
	lastName = Trim(rsUser("lastName"))
	email = Trim(rsUser("email"))
	pswrd = Trim(rsUser("pswrd"))
	signature = Trim(rsUser("signature"))
	noImages = TRIM(rsUser("noImages"))
	qualifications= TRIM(rsUser("qualifications"))
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
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->

<form action="<%= Request.ServerVariables("script_name") %>" method="post" onSubmit="return isReady(this)";>
	<input type="hidden" name="IDuser" value="<%= userID %>">
	
	<h1>Edit User</h1>
	
	<div class="two columns alpha">Date Entered</div>
	<div class="one columns"><%= dateEntered %></div>
	<div class="two columns">View Images</div>
	<div class="one columns">
		<input type="radio" name="noImages" value="0"<% IF noImages=0 THEN %> checked<% END IF%>>Yes
		<input type="radio" name="noImages" value="1"<% IF noImages=1 THEN %> checked<% END IF%>>No
	</div>
	<div class="three columns"><input type="submit" value="Update User"></div>
	<div class="three columns omega"><input type="button" value="Delete User" onClick="location='deleteUser.asp?IDuser=<%= Request("IDuser") %>'; return false";></div>
	<div class="cleaner"></div>
	
	<div class="two columns alpha">First Name</div>
	<div class="four columns"><input type="text" name="firstName" value="<%= firstName %>"></div>
	<div class="two columns">Last Name</div>
	<div class="four columns omega"><input type="text" name="lastName"	value="<%= lastName %>"></div>

	<div class="two columns alpha">Email</div>
	<div class="ten columns omega"><input type="text" name="email" value="<%= email %>"></div>
	
<% If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
	<div class="two columns alpha">Password</div>
	<div class="eight columns"><input type="password" name="pswrd" value="<%= pswrd %>"></div>
	<div class="two columns omega"><input type="button" value="View" onClick="alert('Password: ' + form.pswrd.value)";></div>
	
	<div class="two columns alpha">Signature File</div>
	<div class="six columns">
		<select name="signature">
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
		</select>
	</div>
	<div class="four columns omega"><input type="button" value="Upload Signature File" onClick="location='upSigEditUser.asp?userID=<%= userID %>'; return false";></div>

	<div class="two columns alpha">Qualifications</div>
	<div class="ten columns omega"><textarea cols="50" rows="3" name="qualifications"><%= REPLACE(qualifications,"#@#","'") %></textarea></div>

<% ELSE %>
	<INPUT type="hidden" name="email" value="<%= email %>">
	<INPUT type="hidden" name="pswrd" value="<%= pswrd %>">
	<INPUT type="hidden" name="signature" value="<%= signature %>">
	<INPUT type="hidden" name="noImages" value="<%= noImages %>">
	<INPUT type="hidden" name="qualifications" value="<%= REPLACE(qualifications,"#@#","'") %>">
<% END IF %>

<!--- ----------------------------------------- Rights --------------------------------------- --->
<h3>Rights</h3>
<table width="100%" border="0">		
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
			<th>action</th>	
			<th>erosion</>
<%	End If
 	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>
			<th>director</th>
			<th>current director</th>
			<th>inspector</th></tr>	
<% end if 'Session("validAdmin")
SQL1 = "SELECT p.*, u.userID, u.firstName, u.lastName, u.rights as rights1, pu.rights as rights2" &_
	" FROM Projects as p LEFT JOIN ProjectsUsers as pu ON p.projectID=pu.projectID LEFT JOIN Users as u" &_
	" ON pu.userID=u.userID"
	IF Session("validDirector") AND NOT(Session("validAdmin")) THEN 
		SQL1=SQL1 & " WHERE p.projectID IN (SELECT projectID FROM ProjectsUsers" &_
		" WHERE userID=" & listUserID &" AND rights='director')"
	END IF
SQL1=SQL1 & " ORDER BY projectName ASC, projectPhase ASC"
'--Response.Write(SQL1)
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
			CASE "inspector"	insChecked=True
			CASE "email"		recEmailChecked=True
			CASE "ecc"		    recCCChecked=True
			CASE "bcc"		    recBCCChecked=True
			CASE "action"		actChecked=True
			CASE "erosion"		eroChecked=True
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
<!--- ----------------------------------------- Action ----------------------------------------- --->
		<td align=center><input type="checkbox" name="act<%= compCount %>" value="<%= dispProjID %>"
			<% If actChecked then %>checked<% End If %>></td>		
<!--- ----------------------------------------- Erosion ----------------------------------------- --->
		<td align=center><input type="checkbox" name="ero<%= compCount %>" value="<%= dispProjID %>"
			<% If eroChecked then %>checked<% End If %>></td>
<% 	End If
	If Session("validAdmin") then 'only admin may set rights for other admin, directors and inspectors %>		
<!--- ----------------------------------------- Director --------------------------------------- --->
		<td align=center><input type="checkbox" name="dir<%= compCount %>" value="<%= dispProjID %>"
			<% If dirChecked then %>checked<% End If %>
			<% IF NOT(dirChecked) AND NOT(dirName="None") THEN%>disabled<%END IF%>></td>
<!--- ----------------------------------------- Director Name ---------------------------------- --->
		<td><%= dirName %></td>
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
</div>
</body>
</html>
<%'--	Release Resources ---------------------------------------------------------------------------
RS1.Close
Set RS1 = Nothing %>