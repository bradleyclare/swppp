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

%><!-- #include file="../admin/connSWPPP.asp" --><%

currentDate = date()

If Request.Form.Count > 0 Then
	update = 0
	for n = 0 to 999 step 1
		'Response.Write("coord:coID:" & CStr(n)&":"& Request("coord:coID:" & CStr(n)) &"<br/>")
		if Trim(Request("coord:coID:" & CStr(n))) = "" then
			exit for
		end if
        'Response.Write(Cstr(n) & " s-" & Request("coord:complete:"& CStr(n)) & " n-" & Request("coord:nln:"& CStr(n)) & " ")
        if Request("coord:complete:"& CStr(n)) <> "on" and Request("coord:nln:"& CStr(n)) <> "on" then 
            'Response.Write(" Neither On ")
			coID = Request("coord:coID:"& CStr(n))
            SQLc = "UPDATE Coordinates "& _
			"SET status=0, NLN=0" & _ 
			"WHERE coID = " & coID & ";"
			'Response.Write(SQLc)
			connSWPPP.execute(SQLc)
            
            'update completed item count
            inspecID = Request("coord:inspecID:"& CStr(n))
			SQL1 = "SELECT completedItems, totalItems from Inspections WHERE inspecID = " & inspecID
            Set RS1 = connSWPPP.Execute(SQL1)
            if not RS1.EOF Then
                compItems = RS1("completedItems")
                totItems = RS1("totalItems")
                newItems = compItems - 1
                If newItems < 0 Then
                    completedItems = 0
                Else
                    completedItems = newItems
                End If
            Else
                completedItems = 0
            End If
            
            inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
			" completedItems = " & completedItems & _
			" WHERE inspecID = " & inspecID
		    'response.Write(inspectSQLUPDATE2)
		    connSWPPP.Execute(inspectSQLUPDATE2)

            'add comment to keep track of the status change of the item
            userID  = Session("userID")
            comment = "This item was marked incomplete"
            'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
            SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date)" &_
            " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& currentDate & "')"   
            'response.Write(SQL3)
            Set RS3=connSWPPP.execute(SQL3)
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
<script type="text/javascript">
    function displayCommentWindow(obj) {
        var parts = obj.name.split(":");
        var num = parts[2];

        //display the select div
        var s1 = document.getElementsByName("commentPopup");
        s1[0].className = "commentPopup show";

        //set the hidden div in the select div to remember what number we are modifying
        var s2 = document.getElementsByName("commentIDNum");
        s2[0].value = num;
    }

    function close_popup() {
        //hide the select div
        var s0 = document.getElementsByName("commentPopup");
        s0[0].className = "commentPopup hide";

        //set the num back to -1 so we do not save note
        var s2 = document.getElementsByName("commentIDNum");
        s2[0].value = "-1";
    }

    function save_note(){
        //hide the select div
        var s0 = document.getElementsByName("commentPopup");
        s0[0].className = "commentPopup hide";

        //submit form
        document.getElementById("theForm").submit();
    }
</script>
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
<font size="+1"><b>Completed Items for<br> <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
<br /><br />
<a href="openActionItems.asp?pID=<%=projectID%>&inspecID=<%=inspecID%>">see Open Items</a>
<br /><br />
</center>

<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")&"?pID="&projectID%>" onsubmit="return isReady(this)";>
<table cellpadding="2" cellspacing="0" border="0" width="100%">
	<tr>
        <% If Session("validAdmin") Then %>
            <th width="5%" align="left">Complete</th>
            <th width="5%" align="left">NLN</th>
        <% End If %>
        <th width="5%" align="left">Repeat</th>
        <th width="10%" align="left">ID</th>
        <th width="10%" align="left">Completion Date</th>
        <th width="5%" align="left">Report Date</th>
        <th width="25%" align="left">Location</th>
        <th align="left">Action Item</th>
        <th width="2.5%" align="left">View Note</th>
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
            coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD, NLN" &_
	        " FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
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
                If LD = True Then
                    correctiveMods = "(LD) " & correctiveMods
                End If 
                NLN = rsCoord("NLN")
                If NLN = True Then
                    correctiveMods = "(NLN) " & correctiveMods
                End If 
                commSQLSELECT = "SELECT comment, userID, date" &_
	            " FROM CoordinatesComments WHERE coID=" & coID	
                Set rsComm = connSWPPP.execute(commSQLSELECT)
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
                <td align="left">
                <% If repeat = True Then %>
			        R
		        <% End If %>
                </td>
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
                <% If rsComm.EOF Then %>
                    <td></td>
                <% Else %>
                    <td><button><a href="viewOpenItemComments.asp?coID=<%=coID%>" target="_blank">V</a></button></td>
                <% End If %>
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
    <input type="submit" value="Submit"/><br/><br/>
<% SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& siteMapInspecID
    SET RS3=connSWPPP.execute(SQL3)
    IF NOT(RS3.EOF) THEN 
        sitemap_link = "http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName"))%>
	    <a href='<%=sitemap_link%>'>link for Site Map</a>
    <% END IF %>
</center>
</form>
<br><br>
</body>
</html>

<% connSWPPP.Close
SET connSWPPP=nothing %>
	