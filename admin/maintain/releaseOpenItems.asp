<%Response.Buffer = False%>
<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and not Session("validAdmin") then Response.Redirect("../default.asp") End If
%><!-- #INCLUDE FILE="../connSWPPP.asp" --><%

Server.ScriptTimeout=1500
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
        FOR EACH Item IN Request.Form 'loop through each user
            send_email = True
            currentDate = date()
            strBody=""
            dbgBody=""
            userID = Request(Item)
            SQLSELECT = "SELECT firstName, lastName, email FROM Users WHERE userID = " & userID
            'Response.Write(SQLSELECT & "<br>")
            Set connUsers = connSWPPP.Execute(SQLSELECT) %>
        
            <br />Processing: <%=userID %> - <%=Trim(connUsers("firstName")) %> <%=Trim(connUsers("lastName")) %> - <%=Trim(connUsers("email")) %> -

            <% strBody=strBody &"<head><style>"
            strBody=strBody &"table {border-collapse: collapse;}"
            strBody=strBody &"td{border: 2px solid black; padding: 5px;}"
            strBody=strBody &"th{border: 2px solid black; padding: 5px; font-weight: bold; background-color: grey; color: white;}"
            strBody=strBody &".red{color: #F52006;}"
            strBody=strBody &".green{color: green;}"
            strBody=strBody &".black{color: black;}"
            strBody=strBody &"</style></head>"
            strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
            strBody=strBody &"<center><img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='300'><br><br>"
            strBody=strBody &"<font size='+1'><b>Open Item Report</b></font><br/></center><br/>"
                
            'get all the projects the user is assigned to
            SQLSELECT = "SELECT DISTINCT projectID FROM ProjectsUsers" &_
                " WHERE userID = " & userID &_
                " AND rights <> 'user'"
            'Response.Write(SQLSELECT & "<br>")
            Set connProjUsers = connSWPPP.Execute(SQLSELECT)

            strBody=strBody & "<table>"
            strBody=strBody & "<tr><th>Project Name</th><th>over 5 days</th><th>over 7 days</th><th>over 10 days</th><th>over 14 days</th></tr>"

            'tally up the open items for each project
            'Loop through all projects the user has connection with
            cnt = 0
            Do While Not connProjUsers.EOF
                cnt = cnt + 1
                projID = connProjUsers("projectID")
                dbgBody=dbgBody & projID & "<br/>"

                startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
                endDate=DateAdd("m",1,startDate)
                endDate=DateAdd("d",-1,endDate)
                SQL0 = "SELECT inspecID, inspecDate, reportType," & _
	                " projectID, projectName, projectPhase, released, includeItems, compliance, totalItems, completedItems" & _
	                " FROM Inspections" & _
	                " WHERE projectID = " & projID &_
                    " AND completedItems < totalItems" &_
                    " AND includeItems = 1" &_
                    " AND compliance = 0" &_
                    " ORDER BY projectName"
                '" AND inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
                'Response.Write(SQL0)
                Set RS0 = connSWPPP.Execute(SQL0)

                'Loop through each inspection report and look for open items
                coordCnt = 0
                coordCnt5 = 0
                coordCnt7 = 0
                coordCnt10 = 0
                coordCnt14 = 0
                coordCntLD5 = 0
                coordCntLD7 = 0
                coordCntLD10 = 0
                coordCntLD14 = 0
                displayProj = False
                                       
                If RS0.EOF Then
		            dbgBody=dbgBody & "No Open Items Found<br/>"
	            Else
                    inspecCnt = 0
                    Do While Not RS0.EOF
                        inspecCnt = inspecCnt + 1
                        projName = Trim(RS0("projectName"))
                        projPhase = Trim(RS0("projectPhase"))
                        inspecID = RS0("inspecID")
                        inspecDate = RS0("inspecDate")
                        totalItems = RS0("totalItems")
                        completedItems = RS0("completedItems")

                        dbgBody=dbgBody & inspecDate & "<br/>"
                    
                        'open items on report tally up the open item dates 
                        coordSQLSELECT = "SELECT coID, coordinates, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD FROM Coordinates" &_
	                        " WHERE inspecID=" & inspecID &_
                            " AND status=0" &_
                            " AND infoOnly=0" &_
                            " ORDER BY orderby"	
                        'Response.Write(coordSQLSELECT)
                        Set rsCoord = connSWPPP.execute(coordSQLSELECT)

                        If rsCoord.EOF Then
		                    dbgBody=dbgBody & "No Open Items Found<br/>"
	                    Else
                            Do While Not rsCoord.EOF
                                coordCnt = coordCnt + 1
                                coID = rsCoord("coID")
			                    correctiveMods = Trim(rsCoord("correctiveMods"))
			                    orderby = rsCoord("orderby")
			                    coordinates = Trim(rsCoord("coordinates"))
			                    assignDate = rsCoord("assignDate") 
			                    completeDate = rsCoord("completeDate")
			                    status = rsCoord("status")
			                    repeat = rsCoord("repeat")
			                    useAddress = rsCoord("useAddress")
			                    address = TRIM(rsCoord("address"))
			                    locationName = TRIM(rsCoord("locationName"))
                                infoOnly = rsCoord("infoOnly")
                                LD = rsCoord("LD")
                                If assignDate = "" Then
					                age = 0
				                Else
					                age = datediff("d",assignDate,currentDate) 
				                End If
                                dbgBody=dbgBody & coID &" "& status &" "& age &"<br/>"
                                
                                If age > 14 Then
                                    coordCnt14 = coordCnt14 + 1
                                    displayProj = True
                                    If LD = True Then
                                        coordCntLD14 = coordCntLD14 + 1
                                    End If
                                End If
                                If age > 10 Then
                                    coordCnt10 = coordCnt10 + 1
                                    displayProj = True
                                    If LD = True Then
                                        coordCntLD10 = coordCntLD10 + 1
                                    End If
                                End If
                                If age > 7 Then
                                    coordCnt7 = coordCnt7 + 1
                                    displayProj = True
                                    If LD = True Then
                                        coordCntLD7 = coordCntLD7 + 1
                                    End If
                                End If
                                If age > 5 Then
                                    coordCnt5 = coordCnt5 + 1
                                    displayProj = True
                                    If LD = True Then
                                        coordCntLD5 = coordCntLD5 + 1
                                    End If
                                End If
                                rsCoord.MoveNext
                            LOOP
                            rsCoord.Close
                            SET rsCoord=nothing
                        End If
                        RS0.MoveNext
                    Loop 'RSO
                    RS0.Close
                    SET RS0=nothing
                End If
                connProjUsers.MoveNext
                If inspecCnt > 0 and coordCnt > 0 and displayProj = True Then
                    reportLink = "http://swppp.com/views/openActionItems.asp?pID=" & projID
                    strBody=strBody & "<tr><th><a href='" & reportLink & "' target='_blank'>" & projName &" "& projPhase &"</th><td>"
                    if coordCnt5 > 0 Then
                        strBody=strBody & coordCnt5 & " (" & coordCntLD5 & " LD)"
                    End If
                    strBody=strBody &"</td><td>"
                    If coordCnt7 > 0 Then
                        strBody=strBody & coordCnt7 & " (" & coordCntLD7 & " LD)" 
                    End If
                    strBody=strBody &"</td><td>"
                    If coordCnt10 > 0 Then
                        strBody=strBody & coordCnt10 & " (" & coordCntLD10 & " LD)"
                    End If
                    strBody=strBody & "</td><td>"
                    If coordCnt14 > 0 Then
                        strBody=strBody & coordCnt14 & " (" & coordCntLD14 & " LD)"
                    End If
                    strBody=strBody & "</td></tr>"
		        End If
            Loop 'connProjUsers
            connProjUsers.Close
            SET connProjUsers=nothing
            strBody=strBody & "</table>"
            strBody=strBody &"<br><br><hr>Website: <a href='http://www.swppp.com'>www.swppp.com</a></center></body>" %>

            <% 'send email
            if send_email Then
                fullName = Trim(connUsers("firstName")) & " " & Trim(connUsers("lastName"))
                contentSubject= "Open Item Report for "& fullName &" on "& currentDate
	            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	            Mailer.FromName    = "Don Wims"
	            Mailer.FromAddress = "dwims@swppp.com"
	            Mailer.RemoteHost = "127.0.0.1"
	            Mailer.Subject    = contentSubject
	            Mailer.BodyText = strBody
	            Mailer.ContentType = "text/html"

                '--------this line of code is for testing the smtp server---------------------
                Mailer.AddBCC contentSubject, "dwims@swppp.com"
                Mailer.AddBCC contentSubject, "brad.leishman@gmail.com"
                '--------this line of code is for testing the smtp server---------------------
                
                Mailer.AddRecipient fullName, Trim(connUsers("email"))
		        if not Mailer.SendMail then %>
			        <div class="red">Mail send failure.- </div><%= Mailer.Response %> <br />
                <% Else %>
			        Email Sent <br />
                <% End If
            Else
                Response.Write(strBody)      
                Response.Write(dbgBody)     
            End If 'send_email
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
    Set connUsers = connSWPPP.Execute(SQLSELECT) %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
    <HTML>
        <HEAD>
	        <TITLE>SWPPP INSPECTIONS :: Admin :: Open Item Reports</TITLE>
	        <LINK REL=stylesheet HREF="../../global.css" type="text/css">
        </HEAD>
        <BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	        marginwidth="5" marginheight="5">

            <% If not Request("print") then %> <!-- #INCLUDE FILE="../adminHeader2.inc" --> <% end if %>
            <h1>Users to Receive Open Item Reports</h1>
            <FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
            <div align="center">
            <table width="100%" border="0">
	        <tr><th><b>Count</b></th>
		        <th><b>First Name</b></th>
		        <th><b>Last Name</b></th>
		        <th><b>Send Alert</b></th></tr>
        <%
	        If connUsers.EOF Then
		        Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			        "are currently no users.</i></b></td></tr>")
	        Else
		        altColors="#ffffff"
		
		        Do While Not connUsers.EOF
			        recCount = recCount + 1 %>
	                <tr align="center" bgcolor="<%= altColors %>"> 
		            <td><%= recCount %></td>
		            <td><%= Trim(connUsers("firstName")) %></td>
		            <td><%= Trim(connUsers("lastName")) %></td>
		            <td><INPUT type="checkbox" name="<%= connUsers("userID")%>" value="<%= connUsers("userID")%>"></td></tr>
                    <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			            connUsers.MoveNext
		        Loop
	        End If ' END No Results Found
            %>
            <div align="center"><br /><br />
                To Send Open Item Alerts via Email to all Users assigned to Receive them<br />
                <input type="submit" value="Send Emails"><br />
            </div>
            </FORM>
        </BODY>
    </HTML>
<% END IF
connSWPPP.close
SET connSWPPP=nothing %>