<%@ Language="VBScript" %>
<!-- #include file="../admin/connSWPPP.asp" --> <%

userID = Request("userID")

if len(Request.QueryString("del")) > 0 Then
    id = Trim(Request("id"))

    'delete comment
    commSQLDELETE = "DELETE FROM CoordinatesComments WHERE coID = " & id
    'Response.Write(commSQLSELECT)
    connSWPPP.execute(commSQLDELETE)
End If

If Request.Form.Count > 0 Then	
    endDate=Request("endDate")
    startDate=Request("startDate")
Else
    endDate=CDATE(Date)
    startDate=DateAdd("m",-1,endDate)
End If

'Response.Write(startDate & " - " & endDate)
commSQLSELECT = "SELECT comment, userID, date, coID" &_
    " FROM CoordinatesComments" &_
    " WHERE date BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
    " ORDER BY date DESC"
'Response.Write(commSQLSELECT)
Set rsComm = connSWPPP.execute(commSQLSELECT) 

'get all the projects the user is assigned to
SQLSELECT = "SELECT DISTINCT pu.projectID, p.projectName, p.projectPhase, p.collectionName" &_
    " FROM ProjectsUsers as pu" &_
    " inner join Projects as p" &_
    " on pu.projectID=p.projectID" &_
    " WHERE pu.userID = " & userID &_
    " ORDER BY p.collectionName, p.projectName, p.projectPhase"
'Response.Write(SQLSELECT & "<br>")
Set connProjUsers = connSWPPP.Execute(SQLSELECT) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Recent User Notes</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
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
    
    <h1>Recent Comments for <%=Trim(connUsers("firstName")) %>&nbsp<%=Trim(connUsers("lastName")) %></h1>
    <br /><br />
</center>

<% If rsComm.EOF Then %>
    <h3>No notes from those dates</h3>
<% Else %>
    <form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?userID=<%=userID%>" onsubmit="return isReady(this)";>
        Start Date (MM/DD/YYYY): <input name="startDate" type="text" value="<%=startDate%>" size="8" />  
        End Date (MM/DD/YYYY): <input name="endDate" type="text" value="<%=endDate%>" size="8" />  
        <input name="submit_coord_btn" type="submit" style="font-size: 20px;" value="Submit"/>
    </form>
    <table>
        <tr>
            <th width="5%">Date</th>
            <th width="30%">Note</th>
            <th width="10%">User</th>
            <th width="15%">Project Name</th>
            <th width="20%">Item</th>
            <th width="15%">Location</th>
            <th width="5%">Inspec Date</th>
        </tr>
    <% comment_cnt = 0
    DO WHILE NOT rsComm.EOF 
        coID = rsComm("coID")
        userID = rsComm("userID")
        comment = rsComm("comment")

        If InStr(comment,"This item was marked") <> 1 Then 'returns the position that the string starts 
            'Get user name
            SQLSELECT = "SELECT firstName, lastName FROM Users WHERE userID = " & userID
            'Response.Write(SQLSELECT & "<br>")
            Set connUsers = connSWPPP.Execute(SQLSELECT)
            userName = connUsers("firstName") & " " & connUsers("lastName")

            'get item information
            coordSQLSELECT = "SELECT coID, inspecID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName" &_
	            " FROM Coordinates WHERE coID=" & coID & " AND status=0 AND NLN=0"
            'Response.Write(coordSQLSELECT)
            Set rsCoord = connSWPPP.execute(coordSQLSELECT)
    
            If rsCoord.EOF Then
                note = False
            Else
                correctiveMods = Trim(rsCoord("correctiveMods"))
		        coordinates = Trim(rsCoord("coordinates"))
                useAddress = rsCoord("useAddress")
		        address = TRIM(rsCoord("address"))
		        locationName = TRIM(rsCoord("locationName")) 
                inspecID = rsCoord("inspecID")
            
                'get report name
                inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, i.projectID" & _
		            " FROM Inspections as i, Projects as p" & _
		            " WHERE i.projectID = p.projectID AND inspecID = " & inspecID
                '--Response.Write(inspecSQLSELECT & "<br>")
	            Set rsReport = connSWPPP.execute(inspecSQLSELECT) 
        
                'only display comment if user is affiliated with the same projectID
                show_note = False
                note_projectID = rsReport("projectID")
                if not connProjUsers.EOF Then 
                    DO WHILE NOT connProjUsers.EOF 
                        user_projectID = connProjUsers("projectID")
                        If note_projectID = user_projectID Then
                            show_note = True
                            Exit Do
                        End If 'end projectID ids equal
                        connProjUsers.MoveNext
                    LOOP 
                End If  'end connProjUsers.EOF
                connProjUsers.MoveFirst 'reset for next comment
        
                If show_note Then 
                    comment_cnt = comment_cnt + 1 %> 
                    <tr>
                        <td><%= rsComm("date") %></td>
                        <td><a href='../../views/openActionItems.asp?pID=<%=rsReport("projectID") %>' target="_blank"><%= rsComm("comment")%></a></td>
                        <td><%= userName %></td>
                
                        <td align="left"><%= rsReport("projectName")%>&nbsp;<%= rsReport("projectPhase") %></td>
                        <td><%= correctiveMods %></td>
                        <td>
                        <% if (useAddress) = False Then %>
			                <%=coordinates%>
		                <% Else %>
			                <%=locationName%> (<%=address%>)
		                <% End If %>
                        </td>
                        <td><%= rsReport("inspecDate")%></td>
                    </tr>
                <% End If 'end show_note
            End If 'end rsCoord.EOF
        End If 'end comment status note
        rsComm.MoveNext
    LOOP 
    If comment_cnt = 0 Then
        Response.Write("No Notes Found.")
    End If %>
    </table>
<% 
End If 'end rsComm.EOF
End If 'user exists

rsReport.Close
SET rsReport=nothing
rsCoord.Close
SET rsCoord=nothing
connProjUsers.Close
SET connProjUsers=nothing
rsComm.Close
SET rsComm=nothing
connSWPPP.Close
Set connSWPPP = Nothing

%>
</body>
</html>
