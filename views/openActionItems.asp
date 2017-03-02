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

projectID = Request("pID")

%><!-- #include file="../admin/connSWPPP.asp" --><%

currentDate = date()

If Request.Form.Count > 0 Then
	update = 1
    'check for comment
    commentNum = CInt(Request("commentIDNum"))
    if commentNum > -1 Then
        update = 0
        coID    = Request("coord:coID:" & commentNum)
        comment = Request("commentBox")
        userID  = Session("userID")
        'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
        SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date)" &_
        " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& currentDate & "')"   
        'response.Write(SQL3)
        Set RS3=connSWPPP.execute(SQL3)
    End If

    'If update Then
	    for n = 0 to 999 step 1
		    'Response.Write("coord:coID:" & CStr(n)&":"& Request("coord:coID:" & CStr(n)) &"<br/>")
		    if Trim(Request("coord:coID:" & CStr(n))) = "" then
			    exit for
		    end if
            if Request("coord:complete:"& CStr(n)) = "on" then 
			    SQLc = "UPDATE Coordinates "& _
			    "SET status=1, completeDate='" & Request("coord:date:"& CStr(n))& "' " & _ 
			    "WHERE coID = " & Request("coord:coID:"& CStr(n)) & ";"
			    'Response.Write(SQLc)
			    connSWPPP.execute(SQLc)

                'update completed item count
                inspecID = Request("coord:inspecID:"& CStr(n))
			    SQL1 = "SELECT completedItems from Inspections WHERE inspecID = " & inspecID
                Set RS1 = connSWPPP.Execute(SQL1)
                if not RS1.EOF Then
                    completedItems = RS1("completedItems") + 1
                Else
                    completedItems = 1
                End If
                inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
			    " completedItems = " & completedItems & _
			    " WHERE inspecID = " & inspecID
		        'response.Write(inspectSQLUPDATE2)
		        connSWPPP.Execute(inspectSQLUPDATE2)
		    End If
	    next	
    'End If
End If

SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2)
Set RS2=connSWPPP.execute(SQL2) %>

<html>
<head>
<STYLE>
tr.highlighted {
	cursor:hand;
	background-color:silver
}
</STYLE>
<title>SWPPP INSPECTIONS - Open Items for <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></title>
<link rel="stylesheet" type="text/css" href="../global.css">
<link href="../css/jquery-ui.min.css" rel="stylesheet" type="text/css"/>
<link href="../css/jquery-ui.structure.min.css" rel="stylesheet" type="text/css"/>
<link href="../css/jquery-ui.theme.min.css" rel="stylesheet" type="text/css"/>
<script src="../js/jquery.js" type="text/javascript"></script>
<script src="../js/jquery-ui.min.js" type="text/javascript"></script>
<script>
  $( function() {
    $( ".datepicker" ).datepicker();
  } );

  function check_all_items(obj) {
    for (i=0; i<999; i++){
        var name = "coord:complete:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = 'on';
            s[0].checked = true;
        } else {
            break;
        }
    }
  }

  function uncheck_all_items(obj){
     for (i=0; i<999; i++){
        var name = "coord:complete:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = 'off';
            s[0].checked = false;
        } else {
            break;
        }
     }
  }

  function apply_date_to_all(obj){
     var s = document.getElementsByName("commonDate"); 
     selDate = s[0].value;
     for (i=0; i<999; i++){
        var name = "coord:date:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = selDate;
        } else {
            break;
        }
     }
  }

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
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")& "?pID=" & projectID %>" onsubmit="return isReady(this)";>
    <input type="hidden" name="commentIDNum" value="-1" />
    <div class="commentPopup hide" name="commentPopup">
    <h3>Enter Note:</h3>
    <textarea rows="3" cols="40" name="commentBox"></textarea>
    <br /><br />
    <input type="button" onclick="save_note()" value="Save Note" />
    <input type="button" onclick="close_popup()" value="Cancel" />
    </div>
    <center>
    <img src="../images/color_logo_report.jpg" width="300"><br><br>
    <font size="+1"><b>Open Items for<br/> (<%=projectID %>) <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
    <br /><br />
    <table><tr>
    <td><input type="button" value="Check all Items" onclick="check_all_items(this)" /></td>
    <td><input type="button" value="Un-Check all Items" onclick="uncheck_all_items(this)" /></td>
    <td><input type="text" name="commonDate" class="datepicker" value="<%= currentDate %>" /></td>
    <td><input type="button" value="Apply Date to All" onclick="apply_date_to_all(this)" /></td>
    </tr></table>
    <br/>
    <a href="completedActionItems.asp?pID= <%=projectID%> &inspecID= <%=inspecID%>">see Completed Items</a>
    <br/><br/>
    </center>
    <table cellpadding="2" cellspacing="0" border="0" width="100%">
	    <tr><th width="5%" align="left">Complete</th>
            <th width="5%" align="left">Repeat</th>
            <th width="5%" align="left">ID</th>
            <th width="10%" align="left">Completion Date</th>
            <th width="5%" align="left">Age</th>
            <th width="5%" align="left">Report Date</th>
            <th width="20%" align="left">Location</th>
            <th width="40%" align="left">Action Item</th>
            <th width="2.5%" align="left">Add Note</th>
            <th width="2.5%" align="left">View Note</th>
	    </tr>
    <% If rsInspectInfo.EOF Then
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
		            If infoOnly = True or NLN = True Then
                       do_nothing = 1 
                    Elseif status = false Then 
                        cnt = cnt + 1
                        If cnt = 1 Then
                            siteMapInspecID = inspecID
                        End If
                        commSQLSELECT = "SELECT comment, userID, date" &_
	                    " FROM CoordinatesComments WHERE coID=" & coID	
                        Set rsComm = connSWPPP.execute(commSQLSELECT) %>
		                <input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
                        <input type="hidden" name="coord:inspecID:<%= n %>" value="<%= inspecID %>" />
		                <tr>
		                <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" /></td>
		                <td align="left">
                        <% If repeat = True Then %>
			                R
		                <% End If %>
                        </td>
		                <td align="left"><%= coID %></td>
		                <td align="left"><input class="datepicker" type="text" name="coord:date:<%= n %>" value="<%= currentDate %>"/></td>
		                <td><%= age %> days</td>
		                <td><%= inspecDate %></td>
                        <td>
		                <% if (useAddress) = False Then %>
			                <%=coordinates%>
		                <% Else %>
			                <%=locationName%> (<%=address%>)
		                <% End If %>
		                </td>
		                <td><%= correctiveMods %></td>
                        <td><input type="button" name="coord:note:<%= n %>" value="A" onclick="displayCommentWindow(this)"/></td>
                        <% If rsComm.EOF Then %>
                            <td></td>
                        <% Else %>
                            <td><button><a href="viewOpenItemComments.asp?coID=<%=coID%>" target="_blank">V</a></button></td>
                        <% End If %>
		                </tr>
                        <tr><td colspan="10"></td></tr>
		                <% n = n + 1
                    End If
		            rsCoord.MoveNext 
 	            LOOP 'loop coordinates 
                If start_n <> n Then %>
                    <tr><td colspan="10"><hr /></td></tr>
                <% End If
            End If
            rsInspectInfo.MoveNext
         LOOP 'loop inpection reports
    End If%>
    </table>
    <center>
    <% If cnt = 0 Then %>
        <h3>There are no open items at this time</h3>
    <% End If %>
    <input type="submit" value="Submit" /><br /><br />
    <% SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& siteMapInspecID
    SET RS3=connSWPPP.execute(SQL3)
    IF NOT(RS3.EOF) THEN 
        sitemap_link = "http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName"))%>
	    <a href='<%=sitemap_link%>'>link for Site Map</a>
    <% END IF %>
    </center>
</form>
<br /><br />
</body>
</html>

<% connSWPPP.Close
SET connSWPPP=nothing %>
	