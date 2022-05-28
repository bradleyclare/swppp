<%Response.Buffer = False%>
<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and not Session("validAdmin") then Response.Redirect("../default.asp") End If
%><!-- #INCLUDE FILE="../connSWPPP.asp" --><%

Server.ScriptTimeout=3000

userGroupID = Request("groupNum")

'Response.Write(Request.Form.Count & "<br>")
IF Request.Form.Count > 0 THEN %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
    <HTML>
    <HEAD>
	    <TITLE>SWPPP INSPECTIONS :: Admin :: Sending Project Report</TITLE>
	    <LINK REL=stylesheet HREF="../../global.css" type="text/css">
    </HEAD>
    <BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	    marginwidth="5" marginheight="5">
    <!-- #INCLUDE FILE="../adminHeader2.inc" -->
    <table>
        <%
        testing = false
        FOR EACH Item IN Request.Form 'loop through each user
            If Item = "testing" Then
                If Request(Item) = "on" then 
                    testing = true 
                     %> TESTING ONLY MODE - NO EMAIL WILL BE SENT <% 
                End If
            ElseIf Item <> "userGroup" Then
                send_email = False
                If testing Then
                    debug_msg = True
                Else
                    debug_msg  = False
                End If
                currentDate = date()
                strBody=""
                userID = Request(Item)
                SQLSELECT = "SELECT firstName, lastName, email FROM Users WHERE userID = " & userID & " and active=1 ORDER BY email"
                'Response.Write(SQLSELECT & "<br>")
                Set connUsers = connSWPPP.Execute(SQLSELECT) %>

                    <br />Processing: <%=userID %> - <%=Trim(connUsers("firstName")) %> <%=Trim(connUsers("lastName")) %> - <%=Trim(connUsers("email")) %> -

                    <% strBody=strBody &"<head><style>"
                    strBody=strBody &"table {border-collapse: collapse;}"
                    strBody=strBody &"td{border: 2px solid black; padding: 1px; font-size: 10px; text-align: center;}"
                    strBody=strBody &"th{border: 2px solid black; padding: 3px; font-weight: bold; background-color: #a1a5ad; color: black;}"
                    strBody=strBody &".red{color: #F52006;}"
                    strBody=strBody &".black{color: black;}"
                    strBody=strBody &"</style></head>"
                    strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
                    strBody=strBody &"<img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='175'><br><br>"
                    strBody=strBody &"<font size='+1'><b>Project List Report</b></font><br/><br/>"
                    
                    'get all the projects the user is assigned to
                    SQLSELECT = "SELECT DISTINCT pu.projectID, p.projectName, p.projectPhase, p.collectionName, p.active" &_
                        " FROM ProjectsUsers as pu" &_
                        " inner join Projects as p" &_
                        " on pu.projectID=p.projectID" &_
                        " WHERE pu.userID = " & userID &_
                        " AND p.active=1" & _
                        " ORDER BY p.collectionName, p.projectName, p.projectPhase"
                    'Response.Write(SQLSELECT & "<br>")
                    Set connProjUsers = connSWPPP.Execute(SQLSELECT)

                    strBody=strBody & "<table>"
                    strBody=strBody & "<tr><th>project name</th><th>group name</th></tr>"

                    'tally up the open items for each project
                    'Loop through all projects the user has connection with
                    cnt = 0
                    iterCnt = 0
                    Do While Not connProjUsers.EOF
                        cnt = cnt + 1
                        projID = connProjUsers("projectID")
                        projName = Trim(connProjUsers("projectName"))
                        projPhase = Trim(connProjUsers("projectPhase"))
                        groupName = ""
                        groupNameRaw = connProjUsers("collectionName")
                        
                        strBody=strBody & "<tr><td>" & projName & " " & projPhase & "</td><td>" & groupNameRaw & "</td></tr>"

                        connProjUsers.MoveNext
                    Loop 'connProjUsers
                    connProjUsers.Close
                    SET connProjUsers=nothing
                    
                    strBody=strBody & "</table>" 
                    link = "http://swppp.com/views/viewCommentsUser.asp?userID=" & userID
                    strBody=strBody & "<h3><a href='"& link &"' target='_blank'>view all notes</a></h3>" 

                    'send email
                    If testing Then
                        send_email = false
                    End If
                    if send_email Then
                        fullName = Trim(connUsers("firstName")) & " " & Trim(connUsers("lastName"))
                        contentSubject= "Open Item Report for "& fullName &" on "& currentDate
                        Set Mailer = Server.CreateObject("Persits.MailSender")
                        Mailer.FromName   = "Don Wims"
                        Mailer.From       = "dwims@swppp.com"
                        Mailer.Host       = "127.0.0.1"
                        Mailer.Subject    = contentSubject
                        Mailer.Body       = strBody
                        Mailer.isHTML     = True

                        '--------this line of code is for testing the smtp server---------------------
                        Mailer.AddBCC "dwims@swppp.com", contentSubject
                        Mailer.AddBCC "jwright@swppp.com", contentSubject
                        'Mailer.AddBCC "brad.leishman@gmail.com", contentSubject
                        '--------this line of code is for testing the smtp server---------------------
                    
                        Mailer.AddAddress Trim(connUsers("email")), fullName
                        On Error Resume Next
                        Mailer.Send
                        If Err <> 0 Then %>
                            <div class="red"><%=connUsers("email")%>: Mail send failure.- </div><%= Err.Description %><br />
                        <% Else %>
                            Email Sent <br />
                        <% End If
                    Else %>
                        No Projects: No Email Sent <br />
                    <% End If

                    If debug_msg=True Then
                        Response.Write(strBody)      
                    End If 'send_email
            End If 'Item <> 
        NEXT 'FOR item %>
        <h4>DONE</h4>
    </table>
<%'	Response.End
ELSE
    SQLSELECT = "SELECT userID, firstName, lastName, rights, active, seeScoring, openItemAlerts" &_
	" FROM Users" &_
    " WHERE active=1 AND openItemAlerts = 1 AND seeScoring = 1" &_
	" ORDER BY lastName"
    'Response.Write(SQLSELECT & "<br>")
    Set connUsers = connSWPPP.Execute(SQLSELECT) 
    
    SQLSELECT = "SELECT userGroupID, userGroupName FROM UserGroups ORDER BY userGroupID"
    'Response.Write(SQLSELECT & "<br>")
    Set connGroups = connSWPPP.Execute(SQLSELECT) %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
    <HTML>
        <HEAD>
	        <TITLE>SWPPP INSPECTIONS :: Admin :: Project Reports</TITLE>
	        <LINK REL=stylesheet HREF="../../global.css" type="text/css">
        </HEAD>
        <script type="text/javascript" language="JavaScript1.2">
        function selectUsers(obj) {

            //get selected group name
            nameArray = obj.value.split(" - ");
            groupNum = nameArray[0];

            //redirect page
            window.location.href = "releaseProjectList.asp?groupNum=" + groupNum;
        }
    </script>
        <BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	        marginwidth="5" marginheight="5">

            <% If not Request("print") then %> <!-- #INCLUDE FILE="../adminHeader2.inc" --> <% end if %>
            <h1>send project list alerts</h1>
            <center><h3>to be on this list users must be set to see scoring and receive open item alerts</h3></center>
            <FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
            <div align="center">
                select the users below to send project list alerts via email<br />
                <input type="submit" value="Send Emails"><input type="checkbox" name="testing"/>test mode (do not send email)<br /><br />
                select user group
                <select name="userGroup" onchange="selectUsers(this)">
                    <option value="0 - No Group">0 - No Group</option>
                <% Do While Not connGroups.EOF %>
                    <option value="<%=connGroups("userGroupID")%> - <%=connGroups("userGroupName")%>" 
                    <% If connGroups("userGroupID") = userGroupID Then %> 
                        selected="selected"
                    <% End If %>
                    ><%=connGroups("userGroupID")%> - <%=connGroups("userGroupName")%></option>
                    <% connGroups.MoveNext 
                LOOP %>
                </select><br /><br />
            </div>
            <% if userGroupID > 0 Then
                SQL0 = "SELECT DISTINCT userID, firstName, lastName, userGroupID FROM Users" & _
                    " WHERE userGroupID = '" & userGroupID & "'" & _
                    " ORDER BY lastName"
                'Response.Write(SQL0)
                Set RS0 = connSWPPP.Execute(SQL0) 
            End If  %>
            <table width="100%" border="0">
	        <tr><th><b>count</b></th>
		        <th><b>first name</b></th>
		        <th><b>last name</b></th>
		        <th><b>send alert</b></th></tr>
            <% If connUsers.EOF Then
		        Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			        "are currently no users.</i></b></td></tr>")
	        Else
		        altColors="#ffffff"
		        Do While Not connUsers.EOF 
                    recCount = recCount + 1 
                    If userGroupID > 0 Then
                        If RS0.EOF Then %>
                            <h3>No users in the defined group.</h3>
                        <% Exit Do
                        Else
                            Do While Not RS0.EOF
                                If connUsers("userID") = RS0("userID") Then %>
	                                <tr align="center" bgcolor="<%= altColors %>">
                                    <td><%= recCount %></td>
		                            <td><%= Trim(connUsers("firstName")) %></td>
		                            <td><%= Trim(connUsers("lastName")) %></td>
		                            <td><INPUT type="checkbox" name="<%= connUsers("userID")%>" value="<%= connUsers("userID")%>" checked="checked"></td></tr>
                                    <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
                                End If 
                                RS0.MoveNext
                            Loop 
                            RS0.MoveFirst
                        End If
                    Else %>
	                    <tr align="center" bgcolor="<%= altColors %>">
		                <td><%= recCount %></td>
		                <td><%= Trim(connUsers("firstName")) %></td>
		                <td><%= Trim(connUsers("lastName")) %></td>
		                <td><INPUT type="checkbox" name="<%= connUsers("userID")%>" value="<%= connUsers("userID")%>"></td></tr>
                        <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
                    End If
			        connUsers.MoveNext
		        Loop
	        End If ' END No Results Found 
            %>
            </table>
            </FORM>
        </BODY>
    </HTML>
<% END IF
connSWPPP.close
SET connSWPPP=nothing %>