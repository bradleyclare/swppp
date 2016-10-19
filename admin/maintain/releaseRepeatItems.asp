<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") then Response.Redirect("../default.asp") End If
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
    <%
	FOR EACH Item IN Request.Form
		    '--	Item is ProjectsUsers.projectID &":"& Inspections.inspecID -------------------
		    '--	Request(Item) is Inspections.inspecID ----------------------------------------
		    '-- need to create the Email content ---------------------------------------------
		    strBody=""
        inspecID = Request(Item)

        inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " &_
            "projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " &_
            "compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax, " &_
            "contactEmail, reportType, inches, bmpsInPlace, sediment, " &_
            "narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems" &_
	            " FROM Inspections, Projects, Users" &_
	            " WHERE inspecID = " & inspecID &_
	            " AND Inspections.projectID = Projects.projectID" &_
	            " AND Inspections.userID = Users.userID"
        '--Response.Write("Inspec: "& inspecSQLSELECT &"<br>")
        Set rsInspec = connSWPPP.Execute(inspecSQLSELECT)
        printName = Trim(rsInspec("firstName")) & " " & Trim(rsInspec("lastName"))

        strBody=strBody &"<head><style>"
        strBody=strBody &".red{color: #F52006;}"
        strBody=strBody &".black{color: black;}"
        strBody=strBody &"</style></head>"
        strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
        strBody=strBody &"<center><img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='300'><br><br>"
        strBody=strBody &"<font size='+1'><b>Repeat Item Report</b></font><hr noshade size='1' width='90%'></center>"

        coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName" &_
	        " FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
        'Response.Write(coordSQLSELECT)
        Set rsCoord = connSWPPP.execute(coordSQLSELECT)
    
        strBody=strBody &"<p><table border='0' cellpadding='3' width='100%' cellspacing='0'>"
        strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"
	    If rsCoord.EOF Then
		    'do nothing
	    Else
		    applyScoring = rsInspec("includeItems")
		    currentDate = date()
		    Do While Not rsCoord.EOF
			    coID = rsCoord("coID")
			    correctiveMods = Trim(rsCoord("correctiveMods"))
			    orderby = rsCoord("orderby")
			    coordinates = Trim(rsCoord("coordinates"))
			    existingBMP = Trim(rsCoord("existingBMP")) 
			    assignDate = rsCoord("assignDate") 
			    completeDate = rsCoord("completeDate")
			    status = rsCoord("status")
			    repeat = rsCoord("repeat")
			    useAddress = rsCoord("useAddress")
			    address = TRIM(rsCoord("address"))
			    locationName = TRIM(rsCoord("locationName"))
			    scoring_class = "black"
			    IF applyScoring THEN
				    IF assignDate = "" THEN
					    age = 0
				    ELSE
					    age = datediff("d",assignDate,currentDate) 
				    END IF
				    IF age > 7 THEN
					    scoring_class = "red"
				    END IF
			    END IF
			    IF useAddress THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class = '"& scoring_class &"'>"&  locationName &"<br></td></tr>"
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>address:</b></td>	<td width='80%' align='left' class = '"& scoring_class &"'>"&  address &"<br></td></tr>"
			    ELSE
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class = '"& scoring_class &"'>"&  coordinates &"<br></td></tr>"
			    END IF
			    IF TRIM(rsCoord("existingBMP"))<>"-1" THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>existing BMP:</b></td><td width='80%' align='left'>"&  existingBMP &"<br></td></tr>"
			    END IF
			    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>action needed:</b></td><td width='80%' align='left'>"&  correctiveMods &"</td></tr>"
			    IF applyScoring and repeat THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  age &"<br></td></tr>"
			    END IF
			    strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"  & vbCrLf
			    rsCoord.MoveNext
		    Loop
	    End If ' END No Results Found
    
        SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& inspecID
        SET RS3=connSWPPP.execute(SQL3)
        IF NOT(RS3.EOF) THEN
            strBody=strBody &"<div align='center'><a href='http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName")) &"'>link for Site Map</a></div>"
            strBody=strBody &"<br><div align='center'><a href='http://www.swppp.com'>link to: www.swppp.com</a></div></Body>"
        END IF

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
	    RS1.Open SQL1, connSWPPP

        '--------------------- process mailing -------------------------------------------
	    contentSubject= "Repeat Item Notification"
	    Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	    Mailer.FromName    = "Don Wims"
	    Mailer.FromAddress = "dwims@swppp.com"
	    Mailer.RemoteHost = "127.0.0.1"
	    Mailer.Subject    = contentSubject
	    Mailer.BodyText = strBody & strImages & "<Body>"
	    Mailer.ContentType = "text/html"


        '--------this line of code is for testing the smtp server---------------------
        'Mailer.AddBCC "SWPPP Server testing", "jzuther@gmail.com"
        '--------this line of code is for testing the smtp server---------------------


        '-- build the recipients list ------------------------------------------------
		DO WHILE NOT RS1.EOF
		    curRights = Trim(RS1("rights"))
            if curRights = "email" then
			    Mailer.AddRecipient Trim(RS1("fullName")), Trim(RS1("email"))
			End if
            if curRights = "ecc" then
			    Mailer.AddCC Trim(RS1("fullName")), Trim(RS1("email"))
			End if
            if curRights = "bcc" then
			    Mailer.AddBCC Trim(RS1("fullName")), Trim(RS1("email"))
			End if
			RS1.MoveNext
		LOOP
		if not Mailer.SendMail then %>
			<FONT color="red">Mail send failure.- </FONT><%= Mailer.Response %><br>
<%		else %>
			Emails Sent<BR>
<%		end if
        '--Response.Write(strBody)
	NEXT 'FOR item
'	Response.End
ELSE
    SQL0 = "SELECT i.projectName, i.projectPhase, i.inspecDate, i.inspecID, pu.projectID, i.ReportType, i.released" &_
	" FROM ProjectsUsers pu JOIN Inspections i ON pu.projectID=i.projectID" &_
	" WHERE pu.rights='inspector' AND i.released=0 AND i.userID=pu.userID AND pu.userID="& Session("userID")
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
            <h1>Send Repeat Items via Email</h1>
            <FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
            <div align="center">
            <table border="0" cellpadding=1 cellspacing=1>
	            <tr><th>Project Name|Phase</th><th>Report Date</th><th>Report Type</th><th>send email</th></tr>
            <% 	DO WHILE NOT RS0.EOF %>
	            <tr><td align="left"><%= RS0("projectName")%>&nbsp;<%= RS0("projectPhase") %></td>
		            <td align="left"><%= RS0("inspecDate") %></td>
		            <td align="left"><%= RS0("ReportType") %></td>
		            <td align="center"><INPUT type="checkbox" name="<%= RS0("projectID")%>:<%= RS0("inspecID")%>" value="<%= RS0("inspecID")%>"></td></tr>
            <% 		RS0.MoveNext
	            LOOP
            RS0.Close
            SET RS0=nothing %>
            </table></div>

            <div align="center"><br><br>To Send Repeat Items via Email to all Users assigned<br>
	            to Receive them and release this report <nobr>click..<input type="submit" value="Send Emails"></nobr></div>
            </FORM>
        </BODY>
    </HTML>
<% END IF
connSWPPP.close
SET connSWPPP=nothing %>