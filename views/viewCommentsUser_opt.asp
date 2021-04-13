<%Response.Buffer = False %>
<!-- #include file="../admin/connSWPPP.asp" -->
<%
userID = Request("userID") 
%>

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
    SQLSELECT = "SELECT firstName, lastName FROM Users WHERE userID = " & userID
    'Response.Write(SQLSELECT & "<br>")
    Set connUsers = connSWPPP.Execute(SQLSELECT)
      
    If connUsers.EOF Then %>
         <h3>User ID: <%=userID %> not found.</h3>
    <% Else %>
    
        Recent Comments for <%=Trim(connUsers("firstName")) %>&nbsp<%=Trim(connUsers("lastName")) %>
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
        If connProjUsers.EOF Then %>
            <h3>No Projects Found.</h3>
        <% Else 
            Do While Not connProjUsers.EOF
                projectID = connProjUsers("projectID")
                'Response.Write("ProjectID:" & projectID & "<br/>")
                groupName = ""
                groupNameRaw = connProjUsers("collectionName")

                startDate=CDATE(Month(Date) &"/1/"& Year(Date)) 
                endDate=DateAdd("m",1,startDate)
                endDate=DateAdd("d",-1,endDate)
                         
                inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, includeItems, p.projectName, p.projectPhase" & _
	                " FROM Projects as p, Inspections as i" & _
	                " WHERE inspecID = i.inspecID" &_
	                " AND i.projectID=p.projectID" &_
	                " AND i.projectID="& projectID &_
                    " AND includeItems=1 " &_
                    " AND inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
	                " ORDER BY inspecDate DESC"
                'Response.Write(inspectInfoSQLSELECT & "<br>")
                Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT) 
            
                If not rsInspectInfo.EOF Then
	                Do While Not rsInspectInfo.EOF   
                        inspecID = rsInspectInfo("inspecID")
                        'Response.Write("InspecID:" & inspecID & "<br/>")
                        inspecDate = Trim(rsInspectInfo("inspecDate"))

                        coordSQLSELECT = "SELECT c.coID, c.coordinates, c.existingBMP, c.correctiveMods, c.orderby, c.assignDate, c.completeDate, c.status," &_
                            " c.repeat, c.useAddress, c.address, c.locationName, c.infoOnly, c.LD, c.NLN, c.osc, cc.comment, cc.userID, cc.date" &_ 
                            " FROM Coordinates as c" &_
                            " inner join CoordinatesComments as cc" &_
                            " on c.coID=cc.coID" &_
                            " WHERE c.inspecID=" & inspecID &_
                            " AND c.status<>'true'" &_
                            " AND c.infoOnly<>'true'" &_ 
                            " AND c.NLN<>'true'" &_
                            " ORDER BY c.orderby"

                        'Response.Write(coordSQLSELECT)
                        Set rsCoord = connSWPPP.execute(coordSQLSELECT)
                        start_n = n
	                    Do While Not rsCoord.EOF
                            'Response.Write(coordSQLSELECT)
                            cnt = cnt + 1
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
                            OSC = rsCoord("osc")
                            If LD = True Then
                                correctiveMods = "(LD) " & correctiveMods
                            End If 
                            If OSC = True Then
                                correctiveMods = "(OSC) " & correctiveMods
                            End If 
                            'userName = rsCoord("firstName") & " " & rsCoord("lastName") 
                            userName = "Name"%>

                            <tr>
                                <td><%= rsCoord("date") %></td>
                                <td><a href='../../views/openActionItems.asp?pID=<%=projectID %>' target="_blank"><%= Trim(rsCoord("comment"))%></a></td>
                                <td><%= rsCoord("userID") %></td>
                
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

                        <% rsCoord.MoveNext
                        LOOP
                        rsCoord.Close
                        SET rsCoord=nothing
                    rsInspectInfo.MoveNext
                    LOOP
                End If 'end rsInspectInfo
                rsInspectInfo.Close
                SET rsInspectInfo=nothing
            connProjUsers.MoveNext
            Loop 'connProjUsers
        End If 'no connProjUsers
        connProjUsers.Close
        SET connProjUsers=nothing %>
    </table>
    <% End If
If cnt = 0 Then %>
    <h3>No Comments Found</h3>
<% End If
connSWPPP.close
SET connSWPPP=nothing %>
</BODY>
</HTML>