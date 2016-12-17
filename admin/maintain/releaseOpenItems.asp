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
            show_summary = True
            currentDate = date()
            strBody=""
            userID = Request(Item)
            SQLSELECT = "SELECT firstName, lastName, email FROM Users WHERE userID = " & userID
            'Response.Write(SQLSELECT & "<br>")
            Set connUsers = connSWPPP.Execute(SQLSELECT) %>
        
            <br />Processing: <%=userID %> - <%=Trim(connUsers("firstName")) %> <%=Trim(connUsers("lastName")) %> - <%=Trim(connUsers("email")) %> -

            <% strBody=strBody &"<head><style>"
            strBody=strBody &"table {border-collapse: collapse;}"
            strBody=strBody &"td{border: 1px solid black; padding: 2px;}"
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

            'tally up the open items for each project
            'Loop through all projects the user has connection with
            cnt = 0
            Do While Not connProjUsers.EOF
                cnt = cnt + 1
                projID = connProjUsers("projectID")
                    
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
                If RS0.EOF Then
		            'strBody=strBody & "No Open Items Found<br/>"
	            Else
                    inspecCnt = 0
                    Do While Not RS0.EOF
                        projName = Trim(RS0("projectName"))
                        projPhase = Trim(RS0("projectPhase"))
                        inspecID = RS0("inspecID")
                        inspecDate = RS0("inspecDate")
                        totalItems = RS0("totalItems")
                        completedItems = RS0("completedItems")

                        inspecCnt = inspecCnt + 1
                        if inspecCnt = 1 Then
                            strBody=strBody & "<hr/><h3>Project: " & projName & " " & projPhase & "</h3>"
                        End If
                            
                        'open items on report tally up the open item dates 
                        coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly" &_
	                        " FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
                        'Response.Write(coordSQLSELECT)
                        Set rsCoord = connSWPPP.execute(coordSQLSELECT)

                        If rsCoord.EOF Then
		                    strBody=strBody & "No Open Items Found<br/>"
	                    Else
                            coordCnt = 0
                            coordCnt7 = 0
                            coordCnt10 = 0
                            coordCnt14 = 0
                            coordCnt14g = 0
                            Do While Not rsCoord.EOF
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
                                If assignDate = "" Then
					                age = 0
				                Else
					                age = datediff("d",assignDate,currentDate) 
				                End If
                                itemCnt = itemCnt + 1
                                
                                If infoOnly = True Then
                                    do_nothing = 1 
                                Elseif age > 0 THEN
                                    coordCnt = coordCnt + 1
                                    If show_summary Then
                                        if coordCnt = 1 and inspecCnt = 1 Then
                                            strBody=strBody & "<table>"
                                            strBody=strBody & "<tr><td>Inspection Date</td><td>Less Than 7 days</td><td>Over 7 days</td><td>Over 10 days</td><td>Over 14 days</td></tr>"
                                        End If
                                        If age <= 7 Then
                                            coordCnt7 = coordCnt7 + 1
                                        ElseIf age <= 10 Then
                                            coordCnt10 = coordCnt10 + 1
                                        ElseIf ang <= 14 Then
                                            coordCnt14 = coordCnt14 + 1
                                        Else
                                            coordCnt14g = coordCnt14g + 1
                                        End If
                                    Else
                                        if coordCnt = 1 Then
                                            strBody=strBody &"<h4>Inspection Date: " & inspecDate & "</h4>"
                                            strBody=strBody & "<table>"
                                        End If
                                        If useAddress Then
				                            strBody=strBody &"<tr><td> "&  locationName &" </td>"
				                            strBody=strBody &"<td> "&  address &" </td>"
			                            Else
				                            strBody=strBody &"<tr><td colspan='2'> "&  coordinates &" </td>"
			                            End If
			                            strBody=strBody &"<td> "&  correctiveMods &" </td>"
				                        strBody=strBody &"<td> "&  age &" days old</td></tr>"
                                    End If
				                End If
                                rsCoord.MoveNext
                            LOOP
                            if coordCnt > 0 Then
                                reportLink = "http://swppp.com/views/reportPrint.asp?inspecID=" & inspecID
                                If show_summary Then
                                    strBody=strBody & "<tr><td><a href='" & reportLink & "'>" & inspecDate &"</td><td>"& coordCnt7 &"</td><td>"& coordCnt10 &"</td><td>"& coordCnt14 &"</td><td>"& coordCnt14g &"</td></tr>"
                                    strBody=strBody & "</table>"
                                Else
                                    strBody=strBody & "</table>"
                                    strBody=strBody & "<br/>Complete Report: <a href='" & reportLink & "'>" & reportLink & "</a>"
                                End If
                            End If
                            rsCoord.Close
                            SET rsCoord=nothing
                        End If
                        RS0.MoveNext
                    Loop 'RSO
                    RS0.Close
                    SET RS0=nothing
                End If
                connProjUsers.MoveNext
		    Loop 'connProjUsers
            connProjUsers.Close
            SET connProjUsers=nothing
            strBody=strBody &"<br><br><hr>Website: <a href='http://www.swppp.com'>www.swppp.com</a></center></Body>" %>

            <% 'send email
            if send_email Then
                fullName = Trim(connUsers("firstName")) & " " & Trim(connUsers("lastName"))
                contentSubject= "Open Item Report for "& fullName &" on "& currentDate
	            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	            Mailer.FromName    = "Don Wims"
	            Mailer.FromAddress = "dwims@swppp.com"
	            Mailer.RemoteHost = "127.0.0.1"
	            Mailer.Subject    = contentSubject
	            Mailer.BodyText = strBody & "</body>"
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