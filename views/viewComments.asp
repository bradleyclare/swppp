<%@ Language="VBScript" %>
<% If _
	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
    Not Session("validErosion") And _
	Not Session("validUser") _
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If

projectID = Request("pID")

%><!-- #include file="../admin/connSWPPP.asp" --><%


SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2)
Set RS2=connSWPPP.execute(SQL2) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : View Comments</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<body>
<center>
    <img src="../images/color_logo_report.jpg" width="300"><br><br>
    <font size="+1"><b>Project Notes for<br/> (<%=projectID %>) <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
    <br /><br />
</center>
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
    <% inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
	" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
	" WHERE pu.userID = " & Session("userID") &_
	" AND i.projectID=p.projectID" &_
	" AND i.projectID="& projectID &_
	" ORDER BY inspecDate DESC"
    'Response.Write(inspectInfoSQLSELECT & "<br>")
    Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT) 
            
    If rsInspectInfo.EOF Then
	    Response.Write("<tr><td colspan='10' align='center'><i style='font-size: 15px'>There are no inspection reports found.</i></td></tr>")
    Else
        n = 0
        cnt = 0
        siteMapInspecID = 0
	    Do While Not rsInspectInfo.EOF   
            inspecID = rsInspectInfo("inspecID")
            inspecDate = Trim(rsInspectInfo("inspecDate"))
            includeItems = rsInspectInfo("includeItems")
        
            if includeItems Then
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
                    If NLN = True Then
                        correctiveMods = "(NLN) " & correctiveMods
                    End If 
		            If infoOnly = True Then
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
            End If 'end include items
        rsInspectInfo.MoveNext
        LOOP
    End If 'end rsInspectInfo
    rsInspectInfo.Close
    SET rsInspectInfo=nothing
    connSWPPP.Close
    Set connSWPPP = Nothing %>
</table>
</body>
</html>
