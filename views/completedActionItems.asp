<%@ Language="VBScript" %><%
projectID = Request("pID")

%><!-- #include file="../admin/connSWPPP.asp" --><%

If Request.Form.Count > 0 Then

	update = 0
	for n = 0 to 999 step 1
		'Response.Write("coord:coID:" & CStr(n)&":"& Request("coord:coID:" & CStr(n)) &"<br/>")
		if Trim(Request("coord:coID:" & CStr(n))) = "" then
			exit for
		end if
        if Request("coord:complete:"& CStr(n)) <> "on" then 
			SQLc = "UPDATE Coordinates "& _
			"SET status=0 " & _ 
			"WHERE coID = " & Request("coord:coID:"& CStr(n)) & ";"
			'Response.Write(SQLc)
			connSWPPP.execute(SQLc)
            
            'update completed item count
            inspecID = Request("coord:inspecID:"& CStr(n))
			SQL1 = "SELECT completedItems from Inspections WHERE inspecID = " & inspecID
            'Response.Write(SQL1)
            Set RS1 = connSWPPP.Execute(SQL1)
            completedItems = RS1("completedItems") - 1
            
            inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
			" completedItems = " & completedItems & _
			" WHERE inspecID = " & inspecID
		    'response.Write(inspectSQLUPDATE2)
		    connSWPPP.Execute(inspectSQLUPDATE2)
		End If
	next	
End If

SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2)
SET RS2=connSWPPP.execute(SQL2) %>

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
</head>

<%
inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, p.projectName, p.projectPhase, ImageCount = (Select Count(ImageID) From Images Where inspecID = i.inspecID)" & _
		" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
		" WHERE pu.userID = " & Session("userID") &_
		" AND i.projectID=p.projectID" &_
		" AND i.projectID="& projectID &_
		" ORDER BY inspecDate DESC"
'Response.Write(inspectInfoSQLSELECT & "<br>")
Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT)
%>

<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center>
<img src="../images/color_logo_report.jpg" width="300"><br><br>
<font size="+1"><b>Completed Items for<br> <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font><hr noshade size="1" width="90%">
</center>

<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")& "?pID=" & projectID %>" onsubmit="return isReady(this)";>
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="5%" align="left">Complete</th><th width="5%" align="left">Repeat</th><th width="10%" align="left">ID</th><th width="10%" align="left">Completion Date</th><th width="5%" align="left">Report Date</th><th width="25%" align="left">Location</th><th width="45%" align="left">Action Item</th></tr>
<% If rsInspectInfo.EOF Then
	Response.Write("<tr><td colspan='4' align='center'><i style='font-size: 15px'>There are no inspection reports found.</i></td></tr>")
Else
    n = 0
	Do While Not rsInspectInfo.EOF   
        inspecID = rsInspectInfo("inspecID")
        inspecDate = Trim(rsInspectInfo("inspecDate"))
        includeItems = rsInspectInfo("includeItems")
        
        if includeItems Then
            coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD" &_
	        " FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
            Set rsCoord = connSWPPP.execute(coordSQLSELECT)
            currentDate = date()
            start_n = n
	        Do While Not rsCoord.EOF	
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
                If LD = True Then
                    correctiveMods = "(LD) " & correctiveMods
                End If 
		        If infoOnly = True Then
                   do_nothing = 1 
                Elseif status = true Then %>
		        <tr>
                <input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
                <input type="hidden" name="coord:inspecID:<%= n %>" value="<%= inspecID %>" />
		        <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" checked /></td>
                <% If repeat = True Then %>
			        <td align="left"><input type="checkbox" name="coord:repeat:<%= n %>" disabled checked/></td>
		        <% Else %>
			        <td align="left"><input type="checkbox" name="coord:repeat:<%= n %>" disabled /></td>
		        <% End If %>
		        <td align="left"><%= coID %></td>
		        <td align="left"><%= completeDate %></td>
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
<center><input type="submit" value="Submit"/><br/><br/>
<% if Session("seeScoring")=True Then %>
    <a href="openActionItems.asp?pID= <%=projectID%> &inspecID= <%=inspecID%>">See Open Actions Items</a></center>
<% End If %>
</form>
<br><br>
</body>
</html>

<% connSWPPP.Close
SET connSWPPP=nothing %>
	