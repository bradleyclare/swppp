<%Response.Buffer = False %>
<!-- #include file="../admin/connSWPPP.asp" -->
<%
userID = Request("userID") %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Sending Repeat Item Reports</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>
<BODY>
    <center>
    <img src="../images/color_logo_report.jpg" width="300"><br><br>
    <%
    currentDate = date()
    SQLSELECT = "SELECT firstName, lastName, email FROM Users WHERE userID = " & userID & " ORDER BY email"
    'Response.Write(SQLSELECT & "<br>")
    Set connUsers = connSWPPP.Execute(SQLSELECT)
      
    If connUsers.EOF Then %>
         <h3>User ID: <%=userID %> not found.</h3>
    <% Else %>
    
        Recent Comments for <%=Trim(connUsers("firstName")) %>  <%=Trim(connUsers("lastName")) %>
        <br /><br />
        </center>

        <% 'get all the projects the user is assigned to
        SQLSELECT = "SELECT DISTINCT pu.projectID, p.projectName, p.projectPhase, p.collectionName" &_
            " FROM ProjectsUsers as pu" &_
            " inner join Projects as p" &_
            " on pu.projectID=p.projectID" &_
            " WHERE pu.userID = " & userID &_
            " ORDER BY p.collectionName, p.projectName, p.projectPhase"
        'Response.Write(SQLSELECT & "<br>")
        Set connProjUsers = connSWPPP.Execute(SQLSELECT) %>
    
        <table>
        <tr>
            <th width="5%">Date</th>
            <th width="25%">Note</th>
            <th width="10%">User</th>
            <th width="15%">Project Name</th>
            <th width="25%">Item</th>
            <th width="15%">Location</th>
            <th width="5%">Inspec Date</th>
        </tr>

        <%
        'tally up the comments for each project
        'Loop through all projects the user has connection with
        cnt = 0
        Do While Not connProjUsers.EOF
            cnt = cnt + 1
            projectID = connProjUsers("projectID")
            groupName = ""
            groupNameRaw = connProjUsers("collectionName")

            startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
            endDate=DateAdd("m",1,startDate)
            endDate=DateAdd("d",-1,endDate)
                         
            inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
	            " FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
	            " WHERE pu.userID = " & userID &_
	            " AND i.projectID=p.projectID" &_
	            " AND i.projectID="& projectID &_
                " AND includeItems=1 " &_
                " AND inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
	            " ORDER BY inspecDate DESC"
            'Response.Write(inspectInfoSQLSELECT & "<br>")
            Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT) 
            
            If Not rsInspectInfo.EOF Then 
                n = 0
                cnt = 0
	            Do While Not rsInspectInfo.EOF   
                    inspecID = rsInspectInfo("inspecID")
                    inspecDate = Trim(rsInspectInfo("inspecDate"))

                    coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD, NLN" &_
	                    " FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
                    Set rsCoord = connSWPPP.execute(coordSQLSELECT)
                    start_n = n
	                Do While Not rsCoord.EOF	
	                    coID = rsCoord("coID")
		                correctiveMods = Trim(rsCoord("correctiveMods"))
		                coordinates = Trim(rsCoord("coordinates"))
		                assignDate = rsCoord("assignDate")
		                if assignDate = "" Then
			                age = "?"
		                Else
			                age = datediff("d",assignDate,currentDate) 
		                End If
		                status = rsCoord("status")
		                repeat = rsCoord("repeat")
		                useAddress = rsCoord("useAddress")
		                address = TRIM(rsCoord("address"))
		                locationName = TRIM(rsCoord("locationName"))
                        infoOnly = rsCoord("infoOnly")
                        LD = rsCoord("LD")
                        If LD = True Then
                            correctiveMods = "(LD) " & correctiveMods
                        End If 
                        NLN = rsCoord("NLN")
		                If infoOnly = True or NLN = True Then
                            do_nothing = 1 
                        Elseif status = false Then 
                            cnt = cnt + 1
                            If cnt = 1 Then
                                siteMapInspecID = inspecID
                            End If
                            commSQLSELECT = "SELECT comment, userID, date" &_
	                        " FROM CoordinatesComments WHERE coID=" & coID	
                            Set rsComm = connSWPPP.execute(commSQLSELECT)
                        
                            If not rsComm.EOF Then
                                DO WHILE NOT rsComm.EOF 
                                    userID = rsComm("userID")
                                    SQLSELECT = "SELECT firstName, lastName FROM Users WHERE userID = " & userID
                                    'Response.Write(SQLSELECT & "<br>")
                                    Set connUsers = connSWPPP.Execute(SQLSELECT)
                                    userName = connUsers("firstName") & " " & connUsers("lastName")%> 

                                    <tr>
                                        <td><%= rsComm("date") %></td>
                                        <td><a href='../../views/openActionItems.asp?pID=<%=projectID %>' target="_blank"><%= rsComm("comment")%></a></td>
                                        <td><%= userName %></td>
                
                                        <td align="left"><%= rsInspectInfo("projectName")%>&nbsp;<%= rsInspectInfo("projectPhase") %></td>
                                        <td><%= correctiveMods %></td>
                                        <td>
                                        <% if (useAddress) = False Then %>
			                                <%=coordinates%>
		                                <% Else %>
			                                <%=locationName%> (<%=address%>)
		                                <% End If %>
                                        </td>
                                        <td><%=inspecDate%></td>
                                    </tr>
                                <% rsComm.MoveNext
                                LOOP 
                                rsComm.Close
                                SET rsComm=nothing
                            End If
                        End If 'end status
                    rsCoord.MoveNext
                    LOOP
                    rsCoord.Close
                    SET rsCoord=nothing
                rsInspectInfo.MoveNext
                LOOP
            End If 'end rsInspectInfo
            rsInspectInfo.Close
            SET rsInspectInfo=nothing
        Loop 'connProjUsers
        connProjUsers.Close
        SET connProjUsers=nothing %>
    </table>
    <% End If
connSWPPP.close
SET connSWPPP=nothing %>
</BODY>
</HTML>