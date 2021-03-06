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
	    <TITLE>SWPPP INSPECTIONS :: Admin :: Sending Repeat Item Reports</TITLE>
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
                SQLSELECT = "SELECT firstName, lastName, email FROM Users WHERE userID = " & userID & " ORDER BY email"
                'Response.Write(SQLSELECT & "<br>")
                Set connUsers = connSWPPP.Execute(SQLSELECT) %>
                <br />
                Processing: <%=userID %> - <%=Trim(connUsers("firstName")) %> <%=Trim(connUsers("lastName")) %> - <%=Trim(connUsers("email")) %> -

                <% strBody=strBody &"<head><style>"
                strBody=strBody &"table {border-collapse: collapse;}"
                strBody=strBody &"td{border: 2px solid black; padding: 1px; font-size: 10px; text-align: center;}"
                strBody=strBody &"th{border: 2px solid black; padding: 3px; font-weight: bold; background-color: #a1a5ad; color: black;}"
                strBody=strBody &".red{color: #F52006;}"
                strBody=strBody &".black{color: black;}"
                strBody=strBody &"</style></head>"
                strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
                strBody=strBody &"<img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='175'><br><br>"
                strBody=strBody &"<font size='+1'><b>Open Item Report</b></font><br/><br/>"
                strBody=strBody &"<b>For a complete list of open items, select the project name and log in.</b><br/>"
                   
                'get all the projects the user is assigned to
                SQLSELECT = "SELECT DISTINCT pu.projectID, p.projectName, p.projectPhase, p.collectionName" &_
                    " FROM ProjectsUsers as pu" &_
                    " inner join Projects as p" &_
                    " on pu.projectID=p.projectID" &_
                    " WHERE pu.userID = " & userID &_
                    " ORDER BY p.collectionName, p.projectName, p.projectPhase"

                'Response.Write(SQLSELECT & "<br>")
                Set connProjUsers = connSWPPP.Execute(SQLSELECT)

                strBody=strBody & "<table>"
                strBody=strBody & "<tr><th>project name</th><th>group name</th><th>over 1 day</th><th>over 5 days</th><th class='red'>over 7 days</th><th class='red'>over 10 days</th><th class='red'>over 14 days</th><th class='red'>repeats</th><th>notes</th><th>alert</th></tr>"

                'tally up the open items for each project
                'Loop through all projects the user has connection with
                cnt = 0
                iterCnt = 0
                Do While Not connProjUsers.EOF
                    cnt = cnt + 1
                    projID = connProjUsers("projectID")
                    groupName = ""
                    groupNameRaw = connProjUsers("collectionName")
                    'Response.Write(groupNameRaw)
                    If debug_msg=True Then
                        Response.Write("ProjID: " & projID & "<br/>")
                    End If
                    startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
                    endDate=DateAdd("m",1,startDate)
                    endDate=DateAdd("d",-1,endDate)
                    SQL0 = "SELECT inspecID, inspecDate, reportType," & _
	                    " projectID, projectName, projectPhase, released, includeItems, compliance, totalItems, completedItems, systemic" & _
	                    " FROM Inspections" & _
	                    " WHERE projectID = " & projID &_
                        " AND includeItems = 1" &_
                        " AND compliance = 0" &_
                        " AND openItemAlert = 1" &_
                        " AND completedItems < totalItems"
                    'Response.Write(SQL0)
                    Set RS0 = connSWPPP.Execute(SQL0)

                    'Loop through each inspection report and look for open items
                    coordCnt = 0
                    coordCnt1 = 0
                    coordCnt5 = 0
                    coordCnt7 = 0
                    coordCnt10 = 0
                    coordCnt14 = 0
                    coordCntLD1 = 0
                    coordCntLD5 = 0
                    coordCntLD7 = 0
                    coordCntLD10 = 0
                    coordCntLD14 = 0
                    repeatCnt = 0
                    repeatCntLD = 0
                    displayProj = False
                    displayComments = False
                    displaySystemic = False
                                       
                    If RS0.EOF Then
                        If debug_msg=True Then
		                     Response.Write("No Open Items Found<br/>")
	                     End If 
                   Else
                       inspecCnt = 0
                       Do While Not RS0.EOF
                            inspecCnt = inspecCnt + 1
                            projName = Trim(RS0("projectName"))
                            projPhase = Trim(RS0("projectPhase"))
                            If groupNameRaw <> "" Then
                                groupName = groupNameRaw
                            End If
                            inspecID = RS0("inspecID")
                            inspecDate = RS0("inspecDate")
                            totalItems = RS0("totalItems")
                            completedItems = RS0("completedItems")
                            If debug_msg=True Then
                                Response.Write(projName & ": " & projPhase & ": " & inspecDate & ", total: " & totalItems & ", completed: " & completedItems &"<br/>")
                            End If

                            if completedItems < totalItems then
                               'open items on report tally up the open item dates 
                               coordSQLSELECT = "SELECT coID, assignDate, status, repeat, NLN, LD FROM Coordinates" &_
	                               " WHERE inspecID=" & inspecID &_
                                   " AND status=0" &_
                                   " AND infoOnly=0" &_
                                   " ORDER BY orderby"	
                               'Response.Write(coordSQLSELECT)
                               Set rsCoord = connSWPPP.execute(coordSQLSELECT)

                               If rsCoord.EOF Then
		                           'no nothing
	                            Else
                                   Do While Not rsCoord.EOF
                                       iterCnt = iterCnt + 1
                                       coordCnt = coordCnt + 1
                                       coID = rsCoord("coID")
			                              assignDate = rsCoord("assignDate") 
                                       status = rsCoord("status")
                                       repeat = rsCoord("repeat")
                                       NLN = rsCoord("NLN")
                                       LD = rsCoord("LD")
                                       If assignDate = "" Then
					                           age = 0
				                           Else
					                           age = datediff("d",assignDate,currentDate) 
				                           End If
                                       If debug_msg=True Then
                                          Response.Write("ID: " & coID &", Age: "& age &", Status: "& status &", LD: "& LD &", Repeat: "& repeat & ", Systemic: " & RS0("systemic") &"<br/>")
                                	    End If
                                       
                                       If NLN = True Then
                                          'continue
	                                    ElseIf repeat = True Then
                                          displayProj = True
                                          repeatCnt = repeatCnt + 1
                                          if LD = True Then
                                             repeatCntLD = repeatCntLD + 1
                                          End If
                                       Else
                                          If age > 14 Then
	                                             coordCnt14 = coordCnt14 + 1
	                                             displayProj = True
	                                             If LD = True Then
	                                                coordCntLD14 = coordCntLD14 + 1
	                                             End If
	                                       ElseIf age > 10 Then
	                                             coordCnt10 = coordCnt10 + 1
	                                             displayProj = True
	                                             If LD = True Then
	                                                coordCntLD10 = coordCntLD10 + 1
	                                             End If
	                                       ElseIf age > 7 Then
	                                             coordCnt7 = coordCnt7 + 1
	                                             displayProj = True
	                                             If LD = True Then
	                                                coordCntLD7 = coordCntLD7 + 1
	                                             End If
	                                       ElseIf age > 5 Then
	                                             coordCnt5 = coordCnt5 + 1
	                                             displayProj = True
	                                             If LD = True Then
	                                                coordCntLD5 = coordCntLD5 + 1
	                                             End If
	                                       ElseIf age > 1 Then
	                                             coordCnt1 = coordCnt1 + 1
	                                             displayProj = True
	                                             If LD = True Then
	                                                coordCntLD1 = coordCntLD1 + 1
	                                             End If
	                                       End If
                                       End If 'end repeat

	                                    'check for comments
	                                    commSQLSELECT = "SELECT comment, userID, date" &_
		                                    " FROM CoordinatesComments WHERE coID=" & coID	
	                                    Set rsComm = connSWPPP.execute(commSQLSELECT)       
	                                    if not rsComm.EOF Then
	                                          comment = rsComm("comment")   
                                             if InStr(comment,"This item was marked") <> 1 Then
	                                             displayComments = True
	                                          End If
	                                    End If
                                       
                                       if RS0("systemic") then
                                          displaySystemic = True
                                       Else
                                          displaySystemic = False
                                       end if
                                      rsCoord.MoveNext
                                   LOOP
                                   rsCoord.Close
                                   SET rsCoord=nothing
                               End If
                            End If
                            RS0.MoveNext
                        Loop 'RSO
                        RS0.Close
                        SET RS0=nothing
                    End If
                    connProjUsers.MoveNext
                    If debug_msg=True Then
                        Response.Write("inspecCnt: " & inspecCnt & ", coordCnt: " & coordCnt & ", displayProj: " & displayProj & "<br/>")
                    End If 
                    If inspecCnt > 0 and coordCnt > 0 and displayProj = True Then
                        reportLink = "http://swppp.com/views/openActionItems.asp?pID=" & projID
                        strBody=strBody & VBCrLf & "<tr><td><a href='" & reportLink & "' target='_blank'>" & projName &" "& projPhase &"</td><td>"& groupName &"</td><td>"
                        If coordCnt1 > 0 Then
                            send_email = True
                            strBody=strBody & coordCnt1
                            If coordCntLD1 > 0 Then
                                strBody=strBody & " (" & coordCntLD1 & " LD)"
                            End If 
                        End If
                        strBody=strBody &"</td><td>"
                        If coordCnt5 > 0 Then
                            send_email = True
                            strBody=strBody & coordCnt5
                            If coordCntLD5 > 0 Then
                                strBody=strBody & " (" & coordCntLD5 & " LD)"
                            End If 
                        End If
                        strBody=strBody &"</td><td class='red'>"
                        If coordCnt7 > 0 Then
                            send_email = True
                            strBody=strBody & coordCnt7
                            If coordCntLD7 > 0 Then
                                strBody=strBody & " (" & coordCntLD7 & " LD)"
                            End If 
                        End If
                        strBody=strBody &"</td><td class='red'>"
                        If coordCnt10 > 0 Then
                            send_email = True
                            strBody=strBody & coordCnt10
                            If coordCntLD10 > 0 Then
                                strBody=strBody & " (" & coordCntLD10 & " LD)"
                            End If 
                        End If
                        strBody=strBody & "</td><td class='red'>"
                        If coordCnt14 > 0 Then
                            send_email = True
                            strBody=strBody & coordCnt14
                            If coordCntLD14 > 0 Then
                                strBody=strBody & " (" & coordCntLD14 & " LD)"
                            End If 
                        End If
                        strBody=strBody & "</td><td class='red'>"
                        If repeatCnt > 0 Then
                            send_email = True
                            strBody=strBody & repeatCnt
                            If repeatCntLD > 0 Then
                                strBody=strBody & " (" & repeatCntLD & " LD)"
                            End If 
                        End If
                        strBody=strBody & "</td><td>"
                        if displayComments Then
                            strBody=strBody & "<a href='http://swppp.com/views/viewComments.asp?pID=" & projID &"'> N </a>"
                        End If
                        
                        strBody=strBody & "</td><td>"
                        if displaySystemic Then
                            strBody=strBody & "<a href='http://swppp.com/views/viewSystemicNote.asp?pID=" & projID &"'> A </a>"
                        End If

                        strBody=strBody & "</td></tr>"
                        If debug_msg=True Then
                           Response.Write("coordCnt1: " & coordCnt1 & ", coordCnt5: " & coordCnt5 & ", coordCnt7: " & coordCnt7 & ", coordCnt10: " & coordCnt10 & ", coordCnt14: " & coordCnt14 & ", repeatCnt: " & repeatCnt &", iterCnt: " & iterCnt &", sendEmail: " & send_email & "<br/>")   
                        End If
		            End If
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
                    'Mailer.AddBCC "brad.leishman@gmail.com", contentSubject
                    '--------this line of code is for testing the smtp server---------------------
                
                    Mailer.AddAddress Trim(connUsers("email")), fullName
		            On Error Resume Next
					Mailer.Send
					If Err <> 0 Then %>
			            <div class="red">Mail send failure.- </div><%= Err.Description %> <br />
                    <% Else %>
			            Email Sent <br />
                    <% End If
                Else %>
                    No Open Items: No Email Sent <br />
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
    SQLSELECT = "SELECT userID, firstName, lastName, rights, seeScoring, openItemAlerts" &_
	" FROM Users" &_
    " WHERE openItemAlerts = 1 AND seeScoring = 1" &_
	" ORDER BY lastName"
    'Response.Write(SQLSELECT & "<br>")
    Set connUsers = connSWPPP.Execute(SQLSELECT) 
    
    SQLSELECT = "SELECT userGroupID, userGroupName FROM UserGroups ORDER BY userGroupID"
    'Response.Write(SQLSELECT & "<br>")
    Set connGroups = connSWPPP.Execute(SQLSELECT) %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
    <HTML>
        <HEAD>
	        <TITLE>SWPPP INSPECTIONS :: Admin :: Open Item Reports</TITLE>
	        <LINK REL=stylesheet HREF="../../global.css" type="text/css">
        </HEAD>
        <script type="text/javascript" language="JavaScript1.2">
        function selectUsers(obj) {

            //get selected group name
            nameArray = obj.value.split(" - ");
            groupNum = nameArray[0];

            //redirect page
            window.location.href = "releaseOpenItems.asp?groupNum=" + groupNum;
        }
    </script>
        <BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	        marginwidth="5" marginheight="5">

            <% If not Request("print") then %> <!-- #INCLUDE FILE="../adminHeader2.inc" --> <% end if %>
            <h1>Send Open Item Alerts</h1>
            <center><h3>To be on this list users must be set to See Scoring and receive Open Item Alerts</h3></center>
            <FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
            <div align="center">
                Select the Users Below To Send Open Item Alerts via Email<br />
                <input type="submit" value="Send Emails"><input type="checkbox" name="testing"/>Test Mode (Do Not Send Email)<br /><br />
                Select User Group
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
	        <tr><th><b>Count</b></th>
		        <th><b>First Name</b></th>
		        <th><b>Last Name</b></th>
		        <th><b>Send Alert</b></th></tr>
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