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
    <% clearFromList = False
    If Request("clearList") = "on" Then
        clearFromList = True
    End If
    FOR EACH Item IN Request.Form
		    '--	Item is ProjectsUsers.projectID &":"& Inspections.inspecID -------------------
		    '--	Request(Item) is Inspections.inspecID ----------------------------------------
		    '-- need to create the Email content ---------------------------------------------
		    strBody=""
        inspecID = Request(Item)
        If inspecID = "on" Then
            hello = 1
        Else
            inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " &_
                "projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " &_
                "compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax, " &_
                "contactEmail, reportType, inches, bmpsInPlace, sediment, " &_
                "narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems, sentRepeatItemReport" &_
	                " FROM Inspections, Projects, Users" &_
	                " WHERE inspecID = " & inspecID &_
	                " AND Inspections.projectID = Projects.projectID" &_
	                " AND Inspections.userID = Users.userID"
            'Response.Write("Inspec: "& inspecSQLSELECT &"<br>")
            Set rsInspec = connSWPPP.Execute(inspecSQLSELECT)
            printName = Trim(rsInspec("firstName")) & " " & Trim(rsInspec("lastName"))

            projectName = Trim(rsInspec("projectName"))
            projectPhase = Trim(rsInspec("projectPhase"))
            inspecDate = rsInspec("inspecDate")

            strBody=strBody &"<head><style>"
            strBody=strBody &".red{color: #F52006;}"
            strBody=strBody &".green{color: green;}"
            strBody=strBody &".black{color: black;}"
            strBody=strBody &".ld{font-weight: bold;}"
            strBody=strBody &"</style></head>"
            strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
            strBody=strBody &"<center><img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='300'><br><br>"
            strBody=strBody &"<font size='+1'><b>Repeat Item Report</b></font><br/>"
            strBody=strBody &"<font size='+1'><b>" & projectName & " " & projectPhase & "</b></font><br/>"
            strBody=strBody &"<font size='+1'><b>" & inspecDate & "</center><br/>"

            coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD" &_
	            " FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
            'Response.Write(coordSQLSELECT)
            Set rsCoord = connSWPPP.execute(coordSQLSELECT)
    
            strBody=strBody &"<h3>Repeat Items</h3>"
            strBody=strBody &"<p><table border='0' cellpadding='3' width='100%' cellspacing='0'>"
            strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"
	        If rsCoord.EOF Then
		        'do nothing
	        Else
		        applyScoring = rsInspec("includeItems")
		        currentDate = date()
                send_email = False
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
                    LD = rsCoord("LD")
			        scoring_class = "black"
			        If applyScoring Then
				        If assignDate = "" Then
					        age = 0
				        Else
					        age = datediff("d",assignDate,currentDate) 
				        End If
                        If LD = True Then
                            correctiveMods = "(LD) " & correctiveMods
                            scoring_class = "ld"
                        End If
				        If infoOnly = True Then
                            do_nothing = 1 
                        Elseif  repeat and age > 0 THEN
                            send_email = True
                            scoring_class = "red"
                            If useAddress Then
				                strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class='red'>"&  locationName &"<br></td></tr>"
				                strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>address:</b></td>	<td width='80%' align='left' class='red'>"&  address &"<br></td></tr>"
			                Else
				                strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class='red'>"&  coordinates &"<br></td></tr>"
			                End If
			                strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>action needed:</b></td><td width='80%' align='left' class='red'>"&  correctiveMods &"</td></tr>"
			                If applyScoring Then
				                strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class='red'>"&  age &" days<br></td></tr>"
			                End If
			                strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"  & vbCrLf
				        End If
			        End If
			        rsCoord.MoveNext
		        Loop
	        End If ' END No Results Found
        
            strBody=strBody &"</table><br><center>Complete Report: <a href='http://www.swppp.com/views/reportPrint.asp?inspecID="& inspecID &"'>http://www.swppp.com/views/reportPrint.asp?inspecID="& inspecID &"</a>"
            SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& inspecID
            SET RS3=connSWPPP.execute(SQL3)
            IF NOT(RS3.EOF) THEN
                strBody=strBody &"<br>Site Map: <a href='http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName")) &"'>http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName")) &"</a>"
            END IF
            strBody=strBody &"<br>Website: <a href='http://www.swppp.com'>www.swppp.com</a></center></Body>"

            rsCoord.Close
            Set rsCoord = Nothing
            rsInspec.Close
            Set rsInspec = Nothing
            RS3.Close
            SET RS3 = nothing
    
            '--	now we can create the list of recipients for the email ----------------------------------------
            projID = SPLIT(Item,":")
            projectID = projID(0)
	        '-- Response.Write(Item &":"& Request(Item) &"<br>")
	        SQL1="SELECT DISTINCT (LTRIM(RTRIM(u.firstName)) +' '+ LTRIM(RTRIM(u.lastName))) as fullName,"&_
		        " u.email, u.noImages, i.projectName, i.projectPhase, i.inspecDate, pu.rights" &_
		        " FROM ProjectsUsers pu JOIN Users u on pu.userID=u.userID" &_
		        " JOIN Inspections i ON pu.projectID=i.projectID" &_
		        " WHERE i.inspecID="& Request(Item) &" AND pu.projectID="& projectID
	        Set RS1 = Server.CreateObject("ADODB.Recordset")
	        RS1.Open SQL1, connSWPPP %>

            <tr><td><%=projectName%>&nbsp;<%=projectPhase%></td><td><%=inspecDate%></td>
        
	        <% '--------------------- process mailing ------------------------------------------- 
            updateDB = False
            if clearFromList Then
                updateDB = True
            ElseIf send_email Then
                contentSubject= "Repeat Item Report for "& projectName & " " & projectPhase & " on "& inspecDate
	            Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	            Mailer.FromName    = "Don Wims"
	            Mailer.FromAddress = "dwims@swppp.com"
	            Mailer.RemoteHost = "127.0.0.1"
	            Mailer.Subject    = contentSubject
	            Mailer.BodyText = strBody & strImages & "<Body>"
	            Mailer.ContentType = "text/html"

                '--------this line of code is for testing the smtp server---------------------
                'Mailer.AddBCC "SWPPP Server testing", "brad.leishman@gmail.com"
                '--------this line of code is for testing the smtp server---------------------

                '-- build the recipients list ------------------------------------------------
                prev_email = ""
                DO WHILE NOT RS1.EOF
                    curRights = Trim(RS1("rights"))
                    email     = Trim(RS1("email"))
                    fullname  = Trim(RS1("fullname"))
                    
                    userSQLSELECT = "SELECT userID, pswrd, rights, firstName, lastName, noImages, seeScoring, repeatItemAlerts" &_
		                " FROM Users" & _
		                " WHERE seeScoring = 1 AND repeatItemAlerts = 1 AND email = '" & email & "'"
	                ' Response.Write(userSQLSELECT & "<br>")
	                Set connEmail = connSWPPP.execute(userSQLSELECT)
                    
                    If connEmail.EOF Then
                        'skip user
                    Else    
                        If curRights = "user" or curRights = "email" or curRights = "ecc" or curRights = "bcc" then
			                If prev_email <> email Then
                                prev_email = email
			                    Mailer.AddRecipient fullName, email
			                End If
                        End If
                    End If
			        RS1.MoveNext
		        LOOP
		        if not Mailer.SendMail then %>
			        <td><FONT color="red">Mail send failure.- </FONT><%= Mailer.Response %></td>
        <%		else %>
			        <td>Emails Sent</td>
                    <% updateDB = True
                End If
            Else %>
                <td>No Repeat Items. No Email Sent</td>
        <%  End If 'End send_email 
            'update database to show alert has been sent
            if updateDB Then
                inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
		            " sentRepeatItemReport = 1" & _
		            " WHERE inspecID = " & inspecID
                'response.Write(inspectSQLUPDATE2)
		        connSWPPP.Execute(inspectSQLUPDATE2) %>
                <td> DB Updated </td>
            <% End If %>
            </tr>
            <%'--Response.Write(strBody)
        End If
	NEXT 'FOR item %>
    </table>
<%'	Response.End
ELSE
    startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
    endDate=DateAdd("m",1,startDate)
    endDate=DateAdd("d",-1,endDate)
    SQL0 = "SELECT inspecID, inspecDate, reportType" & _
	", firstName, lastName, i.projectID, i.projectName, i.projectPhase, released, i.sentRepeatItemReport" & _
	" FROM Inspections as i, Users as u, Projects as p" & _
	" WHERE i.userID = u.userID AND i.projectID = p.projectID" &_
	" AND i.inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" 
    If Session("userID") = 42 Then
        SQL0 = SQL0 & " and datediff(m, i.inspecdate, getdate()) <4"
    End	If
    SQL0 = SQL0 & " ORDER BY i.projectName, i.projectPhase, i.inspecDate DESC"
    Set RS0 = Server.CreateObject("ADODB.Recordset")
    'Response.Write(SQL0)
    RS0.Open SQL0, connSWPPP %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
    <HTML>
        <HEAD>
	        <TITLE>SWPPP INSPECTIONS :: Admin :: Repeat Item Reports</TITLE>
	        <LINK REL=stylesheet HREF="../../global.css" type="text/css">
        </HEAD>
        <BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	        marginwidth="5" marginheight="5">

            <% If not Request("print") then %> <!-- #INCLUDE FILE="../adminHeader2.inc" --> <% end if %>
            <h1>Projects with Repeat Items</h1>
            <FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
            <div align="center">
            <table border="1" cellpadding=3px cellspacing=1px>
	            <tr><th>Project Name|Phase</th><th>Inspector</th><th>Report Date</th><th>Report Type</th><th>Send Alert</th></tr>
            <% 	DO WHILE NOT RS0.EOF 
                    inspecID = RS0("inspecID")
                    sentRepeatItemReport = RS0("sentRepeatItemReport")
                    If sentRepeatItemReport = True Then
                        hello = 1
                    Else
                        coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName" &_
	                        " FROM Coordinates WHERE repeat = 1 AND inspecID=" & inspecID 
                        'Response.Write(coordSQLSELECT)
                        Set rsCoord = connSWPPP.execute(coordSQLSELECT)
                        repeatItem = False
                        If rsCoord.EOF Then
                            repeatItem = False
                        Else
                            repeatItem = True
                        End If
                        rsCoord.Close
                        SET rsCoord=nothing
                        IF repeatItem = True THEN %>
	                        <tr><td align="left"><%= RS0("projectName")%>&nbsp;<%= RS0("projectPhase") %></td>
		                    <td align="left"><%= Trim(RS0("firstName"))%>&nbsp;<%=Trim(RS0("lastName"))%></td>
                            <td align="left"><%= RS0("inspecDate") %></td>
		                    <td align="left"><%= RS0("ReportType") %></td>
                            <td align="center"><INPUT type="checkbox" name="<%= RS0("projectID")%>:<%= RS0("inspecID")%>" value="<%= RS0("inspecID")%>" ></td></tr>
                        <% End If 'end has repeatItem
                    End If 'end sentRepeatItemReport		
                    RS0.MoveNext
	            LOOP
            RS0.Close
            SET RS0=nothing %>
            </table></div>

            <div align="center"><br /><br />To Send Repeat Items via Email to all Users assigned<br />
	            to Receive them<br />
                <input type="checkbox" name="clearList" /> Remove From List. Do Not Send Emails<br />
                <input type="submit" value="Send Emails"><br /></div>
            </FORM>
        </BODY>
    </HTML>
<% END IF
connSWPPP.close
SET connSWPPP=nothing %>