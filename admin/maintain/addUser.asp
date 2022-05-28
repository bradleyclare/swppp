<%@ Language="VBScript" %>
<% Response.Buffer = False
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

%> <!-- #include file="../connSWPPP.asp" --> <%
If Request.Form.Count > 0 Then
	highestRights="user"
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	
	Function titleCase(strValue)
		strValue=Ucase(Left(strValue,1)) & Lcase(Mid(strValue,2,Len(strValue)-1))
		titleCase=Replace(strValue,"'","''")
	end function
	
	' is user already in DB?
	userSQLSELECT = "SELECT * FROM Users WHERE email = '" & Request("email") & "'"
	Set connUser = connSWPPP.Execute(userSQLSELECT)
	
	If Not connUser.EOF Then
		userExists = True 
		connUser.Close
		Set connUser = Nothing
		connSWPPP.Close
		Set connSWPPP = Nothing
	Else
		IF IsNull(Request("qualifications")) THEN Request("qualifications")="" END IF
		trimmedQualifications=REPLACE(Request("qualifications"),"'","#@#")
        if Request("seeScoring") = "on" then seeScoring = 1 Else seeScoring = 0 End If
        if Request("openItemAlerts") = "on" then openItemAlerts = 1 Else openItemAlerts = 0 End If
        if Request("repeatItemAlerts") = "on" then repeatItemAlerts = 1 Else repeatItemAlerts = 0 End If
		userSQLINSERT = "INSERT INTO Users (firstName, lastName, email, phone" & _
			", pswrd, dateEntered, signature, noImages, rights, qualifications, seeScoring, openItemAlerts, repeatItemAlerts" & _
			") VALUES (" & _
			"'" & titleCase(Request("firstName")) & "'" & _
			", '" & titleCase(Request("lastName")) & "'" & _
			", '" & strQuoteReplace(Request("email")) & "'" & _
			", '" & strQuoteReplace(Request("phone")) & "'" & _
			", '" & strQuoteReplace(Request("pswrd")) & "'" & _
			", '" & date & "'" & _
			", '" & Request("signature") & "'" & _
			", '" & Request("noImages") & "'" & _
			", 'user'" & _
			", '" & trimmedQualifications & "'" & _
            ", '" & seeScoring & "'" & _
            ", '" & openItemAlerts & "'" & _
            ", '" & repeatItemAlerts & "'" & _
			")"
			
'Response.Write(userSQLINSERT & "<br>")
		connSWPPP.Execute(userSQLINSERT)
		
		maxSQLSELECT = "SELECT MAX(userID) FROM Users"
		Set connUser = connSWPPP.Execute(maxSQLSELECT)
		userID = connUser(0)
		connUser.Close
		Set connUser = Nothing
'-----	rightsValue="0000000" '-- user,action,erosion,email,inspector,director,admin -----------------
		If Request("admin")="on" then rightsValue= "000000001" else rightsValue="000000000" End If
' ----------------------- Inspector, Director, User, Email in Projects User  -------- 
		For Each Item in Request.Form
			Select Case Left(Item,3)
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
		    		If Request("admin")="on" then rightsValue= "000000001" else rightsValue="000000000" End If
					rightsValue=MID(rightsValue,1,4) &"1"& MID(rightsValue,6)
					SQLP="SELECT * FROM Commissions WHERE userID="& userID &" AND projectID="& Request(Item) 
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
					SQL0=SQL0 &" EXEC sp_UpdateCommissions "& userID &", '"& projName &"', "& phase1 &", "& phase2 &", "& phase3 &", "& phase4 &", "& phase5
				Case "dir"
					rights="director"
					rightsValue=MID(rightsValue,1,9) &"1"& MID(rightsValue,11)
			End Select
			If rights<>"" then
				connSWPPP.Execute("sp_InsertPU "& userID &", "& Request(Item) &", '"& rights &"'")
			end if 'item=inspector, director, user or email
			rights=""
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
'		connSWPPP.execute(SQL0)	
		connSWPPP.Close
		Set connSWPPP = Nothing
		Response.Redirect("viewUsersAdmin.asp")
	End If ' is user already in DB?
End If %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Add User</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
	<script language="JavaScript" src="../js/validUsers.js"></script>
	<script language="JavaScript" src="../js/validUsers1.2.js"></script>
</head>

<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr><td><h1>Add User</h1></td></tr></table>
<% If userExists Then %>
	<table width="100%" border="0">
		<tr><td align="center"><font color="red">User email address already exists in database. 
				Go to <a href="viewUsersAdmin.asp">View Users</a> to edit.</font></td></tr></table>
<% Else ' Not userExists %>
<table width="100%" border="0">
	<form action="<% = Request.ServerVariables("script_name") %>" method="post" 
		onSubmit="return isReady(this)";>
		<tr><td width="35%" align="right">first name:</td>
			<td width="65%"><input type="text" name="firstName" size="20" maxlength="20">
			</td></tr>
		<tr><td align="right">last name:</td>
			<td><input type="text" name="lastName" size="20" maxlength="20">
			</td></tr>
		<tr><td align="right">phone:</td>
			<td><input type="text" name="phone" size="15" maxlength="15">
			</td></tr>
		<tr><td align="right">email:</td>
			<td><input type="text" name="email" size="30" maxlength="50"></td></tr>
		<tr><td align="right">password:</td>
			<td><input type="password" name="pswrd" size="15" maxlength="15"></td></tr>
		<tr><td align="right">view images:</td>
			<td><input type="radio" name="noImages" value="0"<% IF noImages=0 THEN %> checked<% END IF%>>Yes
				<input type="radio" name="noImages" value="1"<% IF noImages=1 THEN %> checked<% END IF%>>No</td></tr>
        <tr><td align="right">see scoring:</td>
            <td><input type="checkbox" name="seeScoring" checked /></td>
		</tr>
        <tr><td align="right">receive open item alerts:</td>
            <td><input type="checkbox" name="openItemAlerts" checked /></td>
		</tr>
        <tr><td align="right">receive repeat item alerts:</td>
            <td><input type="checkbox" name="repeatItemAlerts" checked /></td>
		</tr>
<% If Session("validAdmin") then '-- only admin may set inspectors signature files and qualifications %>
		<tr><td align="right">signature file:</td>
			<td><select name="signature">
<% ' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") &"\images\signatures\")
Set gifDirectory = objFolder.Files

For Each gifFile In gifDirectory
	shortenedName = gifFile.Name %>
					<option value="<% = shortenedName %>"
						<% if shortenedName = "dot.gif" Then %> selected<% End If %>><% = shortenedName %>
					</option>
<% Next
Set objFolder = Nothing
Set gifDirectory = Nothing %>
				</select>&nbsp;&nbsp;<input type="button" value="upload signature file" 
				onClick="location='upSigAddUser.asp'; return false";></td></tr>
		<tr><td align="right" valign=top>qualifications:</td>
			<td><TEXTAREA cols="50" rows="3" name="qualifications"></TEXTAREA></td></tr>
<% END IF '-- Valid ADMIN. %>
<!--- ----------------------------------------- Rights --------------------------------------- --->
<table width="100%" border="0">
		<tr><td align="left"><br><font size="+1">Rights</font><br><br></td></tr>
<%  If Session("validAdmin") then 'only admin may set rights for other admin %>		
		<tr><td align="right">Admin:</td>
			<td align=center><input type="checkbox" name="admin"><br></td></tr>
<%	END IF %>
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
			<th>Current Director</th>
			<th>inspector</th></tr>	
<%	end if 'Session("validAdmin") 
SQL1 = "SELECT p.*, u.userID, u.firstName, u.lastName, u.rights as rights1, pu.rights as rights2" &_
	" FROM Projects as p JOIN ProjectsUsers as pu ON p.projectID=pu.projectID LEFT JOIN Users as u" &_
	" ON pu.userID=u.userID"
	IF Session("validDirector") AND NOT(session("validAdmin")) THEN		
		SQL1=SQL1 & " WHERE p.projectID IN (SELECT projectID FROM ProjectsUsers WHERE userID=" & Session("userID") &" AND rights='director') AND active=1"
	END IF
SQL1=SQL1 & " ORDER BY projectName ASC, projectPhase ASC"
SET RS1=connSWPPP.execute(SQL1)
'---	Initialize loop variables -------------------------------------------------------------------
compCount=0 
dispProjID=RS1("projectID")
dispProjName=TRIM(RS1("projectName")) &" "& TRIM(RS1("projectPhase"))
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
				IF TRIM(dirName)="" THEN dirName="None" END IF
				dirChecked=True
		END SELECT
	END IF
'---	If this record does not match the current user, check to find out if it is ------------------
'---	the director for this Project. If it is save director values for it -------------------------
	IF TRIM(RS1("rights2"))="director" THEN dirName= TRIM(RS1("firstName")) &" "& TRIM(RS1("lastName")) END IF
	IF TRIM(dirName)="" THEN dirName="None" END IF
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
			onClick="if (!(document.forms[0].use<%=compCount%>.checked)) { (document.forms[0].emr<%=compCount%>.checked=false); }"></td>
		<td align=center><input type="checkbox" name="emr<%= compCount %>" value="<%= dispProjID %>"
			<% If recEmailChecked then %>checked <% END If%>
			onClick="if (!(document.forms[0].use<%=compCount%>.checked)) { (document.forms[0].emr<%=compCount%>.checked=false); }"></td>
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
			<% IF NOT(dirChecked) AND NOT(dirName="None") THEN%>disabled<%END IF%>></td>
<!--- ----------------------------------------- Director Name ---------------------------------- --->
		<td><%= dirName %></td>
<!--- ----------------------------------------- Inspector -------------------------------------- --->
		<td align=center><input type="checkbox" name="ins<%= compCount %>" value="<%= dispProjID %>"
			<% If insChecked then %>checked<% End If %>></td></tr>		
<% end if 'Session("validAdmin") %>
<%'-- 	Reset the loop Variables --------------------------------------------------------------------
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
		<tr><td colspan=3 align=center><br><input type="submit" value="Add User"></td></tr>
	</form>
</table>
<%	End If ' userExists %>
</body>
</html>