<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

%> <!-- #include file="../connSWPPP.asp" --> <%

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
commSQLSELECT = "SELECT * FROM CoordinatesComments" &_
    " WHERE date BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
    " ORDER BY date DESC, commentID DESC"
'Response.Write(commSQLSELECT)
Set rsComm = connSWPPP.execute(commSQLSELECT) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Recent Notes</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->

<h1>recent notes</h1>
<% If rsComm.EOF Then %>
    <h3>No notes from those dates</h3>
<% Else %>
    <form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
        start date (MM/DD/YYYY): <input name="startDate" type="text" value="<%=startDate%>" size="8" />  
        end date (MM/DD/YYYY): <input name="endDate" type="text" value="<%=endDate%>" size="8" />  
        <input name="submit_coord_btn" type="submit" style="font-size: 20px;" value="submit"/>
    </form>
    <table>
        <tr>
            <th width="5%">date</th>
            <th width="20%">note</th>
            <th width="10%">user</th>
            <th width="15%">project name</th>
            <th width="20%">item</th>
            <th width="15%">location</th>
            <th width="5%">inspec date</th>
            <th width="5%">inspector</th>
            <th width="5%">delete</th>
        </tr>
    <% DO WHILE NOT rsComm.EOF 
        coID = rsComm("coID")
        userID = rsComm("userID")
        comment = rsComm("comment")

        If InStr(comment,"This item was marked") <> 1 Then 'returns the position that the string starts 
            'Get user name
            SQLSELECT = "SELECT firstName, lastName FROM Users WHERE userID = " & userID
            'Response.Write(SQLSELECT & "<br>")
            Set connUsers = connSWPPP.Execute(SQLSELECT)
			if connUsers.EOF Then
				userName = "UNKNOWN"
			else
				userName = connUsers("firstName") & " " & connUsers("lastName")
			end if
			
			'Response.Write(connUsers("firstName") & " " & connUsers("lastName") & "<br>")
			
            'get item information
            coordSQLSELECT = "SELECT * FROM Coordinates WHERE coID=" & coID & " AND status=0 AND NLN=0"
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
				LD = rsCoord("LD")
                OSC = rsCoord("osc")
				
				If LD = True Then
					correctiveMods = "(LD) " & correctiveMods
				End If 
                If OSC = True Then
					correctiveMods = "(OSC) " & correctiveMods
				End If 

                'get report name
                inspecSQLSELECT = "SELECT inspecDate, firstName, lastName, i.projectName, i.projectPhase, i.projectID" & _
		            " FROM Inspections as i, Users as u, Projects as p" & _
		            " WHERE i.userID = u.userID AND i.projectID = p.projectID AND inspecID = " & inspecID
                '--Response.Write(inspecSQLSELECT & "<br>")
	             Set rsReport = connSWPPP.execute(inspecSQLSELECT) %> 

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
                    <td><%= rsReport("firstName") %><%= rsReport("lastName") %></td>
                    <td><a href="recentComments.asp?del=1&id=<%=coID %>"><input type="button" value="Delete" /></a></td>
                </tr>
            <% End If
        End If
        rsComm.MoveNext
    LOOP %>
    </table>
<% End If
rsReport.Close
SET rsReport=nothing
rsCoord.Close
SET rsCoord=nothing
rsComm.Close
SET rsComm=nothing
connSWPPP.Close
Set connSWPPP = Nothing %>
</body>
</html>
