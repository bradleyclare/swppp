<%@ Language="VBScript" %><%

If _
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

projectID = Trim(Request("pID"))
startDate = Trim(Request("startDate"))
endDate = Trim(Request("endDate"))

%><!-- #include file="../admin/connSWPPP.asp" --><%

SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2)
SET RS2=connSWPPP.execute(SQL2) 

SQLH="SELECT inspecID FROM Inspections WHERE horton=1 AND projectID="& projectID
SET RSH=connSWPPP.execute(SQLH)
hortonFlag=False
completePast="completed"
completeAction="complete"
if NOT(RSH.EOF) THEN 
    hortonFlag=True 
    completePast="closed"
    completeAction="close"
END IF
%>

<html>
<head>
<STYLE>
tr.highlighted {
	cursor:hand;
	background-color:silver
}
</STYLE>
<title>SWPPP INSPECTIONS - Completed Items for <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></title>
<link rel="stylesheet" type="text/css" href="../global.css">

<%
inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
		" WHERE i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
      " AND inspecDate BETWEEN '"& startDate &"' AND '"& endDate &"'" &_
		" ORDER BY inspecDate DESC"
'Response.Write(inspectInfoSQLSELECT & "<br>")
Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT)
%>

<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center>
<img src="../images/color_logo_report.jpg" width="300"><br><br>
<font size="+1"><b><%=completePast%> items for<br> <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
<br /><br />
</center>

<table cellpadding="2" cellspacing="0" border="0" width="100%">
	<tr>
        <% If Session("validAdmin") or Session("validDirector") Then %>
            <th width="5%" align="left"><%=completePast%></th>
            <th width="2.5%" align="left">NLN</th>
        <% End If %>
        <th width="5%" align="left">ID</th>
        <th width="10%" align="left"><%=completeAction%> date</th>  
        <th width="5%" align="left">report date</th>
        <th width="15%" align="left">location</th>
        <th align="left">action item</th>
    </tr>
<% If rsInspectInfo.EOF Then
	Response.Write("<tr><td colspan='4' align='center'><i style='font-size: 15px'>There are no inspection reports found.</i></td></tr>")
Else
    n = 0
    siteMapInspecID = 0
	Do While Not rsInspectInfo.EOF   
        inspecID = rsInspectInfo("inspecID")
        inspecDate = Trim(rsInspectInfo("inspecDate"))
        includeItems = rsInspectInfo("includeItems")
        
        if includeItems Then
            coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
            Set rsCoord = connSWPPP.execute(coordSQLSELECT)
            currentDate = date()
            start_n = n
	         Do While Not rsCoord.EOF	
               If n = 0 Then
                  siteMapInspecID = inspecID
               End If
	            coID = rsCoord("coID")
		         correctiveMods = Trim(rsCoord("correctiveMods"))
		         coordinates = Trim(rsCoord("coordinates"))
		         assignDate = rsCoord("assignDate")
		         completeDate = rsCoord("completeDate")
		         if assignDate = "" Then
			         age = "?"
		         Else
			         age = datediff("d",assignDate,completeDate) 
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
               NLN = rsCoord("NLN")
               If NLN = True Then
                  correctiveMods = "(NLN) " & correctiveMods
               End If 
               commSQLSELECT = "SELECT c.comment, c.userID, c.date, u.firstName, u.lastName" &_
	               " FROM CoordinatesComments as c, Users as u WHERE c.userID = u.userID" &_
                  " AND coID=" & coID	
               'Response.Write(commSQLSELECT)
               Set rsComm = connSWPPP.execute(commSQLSELECT)
               completer = ""
               show_note = false
               If rsComm.EOF Then
                  show_note = false
               Else
                  'find the completion note
                  Do While Not rsComm.EOF
                     If rsComm("comment") = "This item was marked complete" or rsComm("comment") = "This item was marked NLN" Then
                        completer = rsComm("firstName") & " " & rsComm("lastName")
                        completeDate = rsComm("date")
                     Else
                        show_note = true
                     End If
                     rsComm.MoveNext
                  LOOP 
               End If
		         If infoOnly = True Then
                  do_nothing = 1 
               Elseif status = true or NLN = true Then %>
		           <tr>
                 <input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
                 <input type="hidden" name="coord:inspecID:<%= n %>" value="<%= inspecID %>" />
                 <% status_str = ""
                 If status = True Then
                   status_str = "checked"
                 End If
                 nln_str = ""
                 If NLN = True Then
                    nln_str = "checked"
                 End If
                 If Session("validAdmin") Then %> 
                    <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" <%=status_str %> /></td>
                    <td align="left"><input type="checkbox" name="coord:nln:<%= n %>" <%=nln_str %> /></td>
                 <% End If %>
		           <td align="left"><%= coID %></td>
		           <td align="left"><%= completeDate %>: <%= completer %></td>
		           <td><%= inspecDate %></td>
                   <td>
		           <% if (useAddress) = False Then %>
			           <%=coordinates%>
		           <% Else %>
			           <%=locationName%> (<%=address%>)
		           <% End If %>
		           </td>
		           <td><%= correctiveMods %></td>
		           </tr>
		           <% n = n + 1
              End If
		        rsCoord.MoveNext
            LOOP 'loop coordinates 
            If start_n <> n Then %>
                <tr><td colspan="8"><hr /></td></tr>
            <% End If
        End If
        rsInspectInfo.MoveNext
     LOOP 'loop inpection reports
End If%>
</table>
<center>
<% If Session("validAdmin") Then  %>
    <input type="submit" value="Submit"/><br/><br/>
<% End If %>
<% SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& siteMapInspecID
SET RS3=connSWPPP.execute(SQL3)
IF NOT(RS3.EOF) THEN 
    sitemap_link = "http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName"))%>
	<a href='<%=sitemap_link%>'>link for Site Map</a>
<% End If %>
</center>
<br><br>
</body>
</html>

<% connSWPPP.Close
SET connSWPPP=nothing %>
	