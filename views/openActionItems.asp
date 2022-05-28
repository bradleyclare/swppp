<%@ Language="VBScript" %><%
Response.Buffer = False
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
allItems = Request("allItems")
locdir = Trim(Request("loc"))

%><!-- #include file="../admin/connSWPPP.asp" --><%

SQL1="SELECT * FROM ProjectsUsers WHERE userID="& Session("userID") & " AND projectID="& projectID
SET RS1=connSWPPP.execute(SQL1)
projectValidUser=False
if NOT(RS1.EOF) THEN projectValidUser=True End If
if Session("validAdmin") THEN projectValidUser=True End If

if projectValidUser=False Then
    Response.Redirect("projects.asp")
End If

Server.ScriptTimeout=1500

currentDate = date()

SQLH="SELECT inspecID FROM Inspections WHERE horton=1 AND projectID="& projectID
SET RSH=connSWPPP.execute(SQLH)
hortonFlag=False
completePast="completed"
completeAction="complete"
completeDate = "completion"
if NOT(RSH.EOF) THEN 
    hortonFlag=True 
    completePast="closed"
    completeAction="close"
    completeDate="close"
END IF

SQLH="SELECT inspecID FROM Inspections WHERE forestar=1 AND projectID="& projectID
SET RSH=connSWPPP.execute(SQLH)
forestarFlag=False
if NOT(RSH.EOF) THEN 
    forestarFlag=True 
    if Session("validErosion") Then
        completePast="completed"
        completeAction="complete"
        completeDate = "completion"
    else
        completePast="closed"
        completeAction="close"
        completeDate = "completion"
    end if
END IF

If Request.Form.Count > 0 Then
	update = 1
    'check for comment
    commentNum = CInt(Request("commentIDNum"))
    if commentNum > -1 Then
        update = 0
        coID    = Request("coord:coID:" & commentNum)
        comment = Replace(Request("commentBox"),"'","''")
        userID  = Session("userID")
		inspecID = Request("coord:inspecID:"& commentNum)
        'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
        SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date, inspecID, projectID)" &_
        " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& currentDate & "', "& inspecID & ", "& projectID & ")"   
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
                'if this is a horton project the following things will happen when an item is marked complete, 
                '1-erosion user - marking the item complete will just add comment to say the item is complete but it won't close the item
                '2-other user - marking item complete will close the item like normal
                coID = Request("coord:coID:"& CStr(n))
                inspecID = Request("coord:inspecID:"& CStr(n))
                If (hortonFlag or forestarFlag) and Session("validErosion") Then
                    'add comment to keep track of the status change of the item
                    userID  = Session("userID")
                    comment = "This item was marked done"
                    If forestarFlag Then
                        udate = Request("coord:date:"& CStr(n))
                    else
                        udate = currentDate 
                    End If
                    'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
                    SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date, inspecID, projectID)" &_
                    " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& udate & "', "& inspecID & ", "& projectID & ")"    
                    'response.Write(SQL3)
                    Set RS3=connSWPPP.execute(SQL3)
                Else
                    'for a forestar project if it hasn't been assigned a completion date use the calendar to create a completion date then use the current date to close the item
                    doneStatus = Request("coord:done:"& CStr(n))
                    If forestarFlag and doneStatus<>"done" Then
                        userID  = Session("userID")
                        comment = "This item was marked done"
                        udate = Request("coord:date:"& CStr(n))
                        'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
                        SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date, inspecID, projectID)" &_
                        " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& udate & "', "& inspecID & ", "& projectID & ")"    
                        'response.Write(SQL3)
                        Set RS3=connSWPPP.execute(SQL3)
                        udate = currentDate 'set it to current date for remaining updates below
                    Else
                        udate = Request("coord:date:"& CStr(n))
                    End If

                    'update status to closed
                    SQLc = "UPDATE Coordinates "& _
                    "SET status=1, completeDate='" & udate & "' " & _ 
                    "WHERE coID = " & coID & ";"
                    'Response.Write(SQLc)
                    connSWPPP.execute(SQLc)

                    'update completed item count
                    SQL1 = "SELECT completedItems, totalItems from Inspections WHERE inspecID = " & inspecID
                    Set RS1 = connSWPPP.Execute(SQL1)
                    if not RS1.EOF Then
                        compItems = RS1("completedItems")
                        totItems = RS1("totalItems")
                        newItems = compItems + 1
                        If newItems > totItems Then
                            completedItems = totItems
                        Else
                            completedItems = newItems
                        End If
                    Else
                        completedItems = 1
                    End If
                    inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
                    " completedItems = " & completedItems & _
                    " WHERE inspecID = " & inspecID
                    'response.Write(inspectSQLUPDATE2)
                    connSWPPP.Execute(inspectSQLUPDATE2)

                    'add comment to keep track of the status change of the item
                    userID  = Session("userID")
                    comment = "This item was marked complete"
                    'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
                    SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date, inspecID, projectID)" &_
                    " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& udate & "', "& inspecID & ", "& projectID & ")"    
                    'response.Write(SQL3)
                    Set RS3=connSWPPP.execute(SQL3)
                End If
		    End If
            If Request("coord:NLN:"& CStr(n)) = "on" Then
                coID = Request("coord:coID:"& CStr(n))
                udate = Request("coord:date:"& CStr(n))
                SQLc = "UPDATE Coordinates "& _
			    "SET NLN=1, completeDate='" & udate & "' " & _ 
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
                    newItems = compItems + 1
                    If newItems > totItems Then
                        completedItems = totItems
                    Else
                        completedItems = newItems
                    End If
                Else
                    completedItems = 1
                End If
                inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
			    " completedItems = " & completedItems & _
			    " WHERE inspecID = " & inspecID
		        'response.Write(inspectSQLUPDATE2)
		        connSWPPP.Execute(inspectSQLUPDATE2)

                'add comment to keep track of the status change of the item
                userID  = Session("userID")
                comment = "This item was marked NLN"
                'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
                SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date, inspecID, projectID)" &_
				" VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& udate & "', "& inspecID & ", "& projectID & ")"  
                'response.Write(SQL3)
                Set RS3=connSWPPP.execute(SQL3)
            End If 
	    next	
    'End If
End If

SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2)
Set RS2=connSWPPP.execute(SQL2) 
%>

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
<script type="text/javascript">
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

  function nln_all_items(obj) {
    for (i=0; i<999; i++){
        var name = "coord:NLN:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = 'on';
            s[0].checked = true;
        } else {
            break;
        }
    }
  }

  function unnln_all_items(obj){
     for (i=0; i<999; i++){
        var name = "coord:NLN:" + i.toString();
        var s = document.getElementsByName(name);
        if (s.length > 0){
            s[0].value = 'off';
            s[0].checked = false;
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
currentDate = date()
if allItems Then
    startDate=DateAdd("yyyy",-20,currentDate)
Else
    startDate=DateAdd("yyyy",-1,currentDate)
End If
SQL0 = "SELECT inspecID, inspecDate, reportType, projectID, projectName, projectPhase, released, " & _
    " includeItems, compliance, totalItems, completedItems, horton" & _
    " FROM Inspections" & _
    " WHERE projectID = " & projectID &_
    " AND includeItems = 1" &_ 
    " AND released = 1" &_
    " AND compliance = 0" &_
    " AND openItemAlert = 1" &_
    " AND completedItems < totalItems" &_
    " AND inspecDate BETWEEN '"& startDate &"' AND '"& currentDate &"'" 
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)
%>

<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")&"?pID="&projectID%>" onsubmit="return isReady(this)";>
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
    <font size="+1"><b>open items for<br/> (<%=projectID %>) <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
    <br /><br />
    <table><tr>
    <td><input type="button" value="check all items" onclick="check_all_items(this)" /></td>
    <% If Session("validAdmin") or (projectValidUser and hortonFlag=False and forestarFlag=False) Then %>
        <td><input type="button" value="un-check all items" onclick="uncheck_all_items(this)" /></td>
        <td><input type="text" name="commonDate" class="datepicker" value="<%= currentDate %>" /></td>
        <td><input type="button" value="apply date to all" onclick="apply_date_to_all(this)" /></td>
    <% End If%>
    <% If Session("validAdmin") Then %>
        <td><input type="button" value="NLN all items" onclick="nln_all_items(this)" /></td>
        <td><input type="button" value="un-NLN all items" onclick="unnln_all_items(this)" /></td>
    <% End If %>
    </tr></table>
    <br/>
    <a href="completedActionItems.asp?pID=<%=projectID%>&inspecID=<%=inspecID%>">see <%=completePast%> items</a>
    </br><a href="openActionItems.asp?pID=<%=projectID%>&allItems=1">see all open items</a>
    <br/><br/>
    </center>
    <table cellpadding="2" cellspacing="0" border="0" width="100%">
	    <tr><th width="5%" align="left"><%=completeAction%></th>
            <% If projectValidUser and (hortonFlag or forestarFlag) Then %>
                <th width="5%" align="left">item status</th>
            <% End If %>
            <% If Session("validAdmin") Then %>
                <th width="5%" align="left">NLN</th>
            <% End If %>
            <th width="5%" align="left">repeat</th>
            <th width="5%" align="left">ID</th>
            <th width="10%" align="left"><%=completeDate%> date</th>
            <th width="5%" align="left">age</th>
            <th width="5%" align="left">report date</th>
            <% If locdir = "asc" then %>
            <th width="20%" align="left"><a href="openActionItems.asp?pID=<%=projectID%>&loc=desc">location</a></th>
            <% Else %>
            <th width="20%" align="left"><a href="openActionItems.asp?pID=<%=projectID%>&loc=asc">location</a></th>
            <% End If %>
            <th align="left">action item</th>
            <th width="2.5%" align="left">add note</th>
            <th width="2.5%" align="left">view note</th>
	    </tr>
    <% siteMapInspecID = 0
    show_debug = false
    dbgBody = ""
    If RS0.EOF Then
		dbgBody=dbgBody & "No Open Items Found<br/>"
	    Response.Write("<tr><td colspan='10' align='center'><i style='font-size: 15px'>There are no inspection reports found.</i></td></tr>")
    Else
       n = 0
       inspecCnt = 0
	    Do While Not RS0.EOF   
		    inspecCnt = inspecCnt + 1
          projName = Trim(RS0("projectName"))
          projPhase = Trim(RS0("projectPhase"))
          If groupNameRaw <> "" Then
             groupName = groupNameRaw
          End If
          inspecID = RS0("inspecID")
          inspecDate = RS0("inspecDate")
          totalItems = RS0("totalItems")
          completedItems = RS0("completedItems")
          'If siteMapInspecID = 0 Then
	          siteMapInspecID = inspecID
	       'End If

          dbgBody=dbgBody & projName & ": " & projPhase & ": " & inspecDate & "<br/>"

        	 'open items on report tally up the open item dates 
			 If locdir = "desc" Then
                coordSQLSELECT = "SELECT * FROM Coordinates" &_
                " WHERE inspecID=" & inspecID &_
                " AND status=0" &_
                " AND infoOnly=0" &_
                " ORDER BY locationName DESC"
            Elseif locdir = "asc" Then
                coordSQLSELECT = "SELECT * FROM Coordinates" &_
                " WHERE inspecID=" & inspecID &_
                " AND status=0" &_
                " AND infoOnly=0" &_
                " ORDER BY locationName ASC"
            Else
                coordSQLSELECT = "SELECT * FROM Coordinates" &_
                " WHERE inspecID=" & inspecID &_
                " AND status=0" &_
                " AND infoOnly=0" &_
                " ORDER BY OrderBy"
            End If 	
            'Response.Write(coordSQLSELECT)
            Set rsCoord = connSWPPP.execute(coordSQLSELECT)
            start_n = n
			If rsCoord.EOF Then
                'no nothing
            Else
		        Do While Not rsCoord.EOF	
				    coordCnt = coordCnt + 1
		            coID = rsCoord("coID")
	                correctiveMods = Trim(rsCoord("correctiveMods"))
	                orderby = rsCoord("orderby")
	                coordinates = Trim(rsCoord("coordinates"))
	                assignDate = rsCoord("assignDate") 
	                completeDate = rsCoord("completeDate")
                    repeat = rsCoord("repeat")
	                useAddress = rsCoord("useAddress")
	                address = TRIM(rsCoord("address"))
	                locationName = TRIM(rsCoord("locationName"))
	                infoOnly = rsCoord("infoOnly")
	                LD = rsCoord("LD")
                    NLN = rsCoord("NLN")
                    OSC = rsCoord("osc")
	                parentID = rsCoord("parentID")
	                If assignDate = "" Then
	                    age = 0
	                Else
	                    age = datediff("d",assignDate,currentDate) 
	                End If
				    dbgBody=dbgBody & "ID: " & coID &", Age: "& age &", LD: "& LD &", NLN: "& NLN & "<br/>"
	                If LD = True Then
	                    correctiveMods = "(LD) " & correctiveMods
	                End If
                    If OSC = True Then
	                    correctiveMods = "(OSC) " & correctiveMods
	                End If 
			        If infoOnly = True or NLN = True Then
	                    do_nothing = 1 
	                Elseif status = false Then 
	                    commSQLSELECT = "SELECT comment, userID, date" &_
		                    " FROM CoordinatesComments WHERE coID=" & coID	
	                    Set rsComm = connSWPPP.execute(commSQLSELECT) %>
			            <input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
	                    <input type="hidden" name="coord:inspecID:<%= n %>" value="<%= inspecID %>" />
			            <tr>
			            <% If projectValidUser and (hortonFlag or forestarFlag) Then 
                        commSQLSELECT = "SELECT c.comment, c.userID, c.date, u.firstName, u.lastName" &_
                            " FROM CoordinatesComments as c, Users as u WHERE c.userID = u.userID" &_
                        " AND coID=" & coID	
                        'Response.Write(commSQLSELECT)
                        Set rsComm2 = connSWPPP.execute(commSQLSELECT)
                        done = ""
                        completer = ""
                        completeDate = ""
                        If not rsComm2.EOF Then
                            'find the completion note
                            Do While Not rsComm2.EOF
                                If rsComm2("comment") = "This item was marked done" Then
                                    done = "done"
                                    completer = rsComm2("firstName") & " " & rsComm2("lastName")
                                    completeDate = rsComm2("date")
                                End If
                                rsComm2.MoveNext
                            LOOP 
                        End If %>
                    <% End If %>
                    <input type="hidden" name="coord:done:<%= n %>" value="<%= done %>" />
                    <% If projectValidUser=False Then %>
                        <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" disabled="disabled" /></td>
                    <% ElseIf (hortonFlag or forestarFlag) and Session("validErosion") and completer <> "" Then %>
                        <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" disabled="disabled" /></td>
                    <% Else %>
                        <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" /></td>
                    <% End If %>
                    <% If projectValidUser and (hortonFlag or forestarFlag) Then %>
                        <% If Session("validAdmin") Then %>
                            <td><a href="viewOpenItemComments.asp?coID=<%=coID%>" target="_blank"><%=done%></br><%=completer%></br><%=completeDate%></a></td>
                        <% Else %>
                            <td><%=done%></br><%=completer%></br><%=completeDate%></td>
                        <% End If %>
                    <% End If %>
                    <% If Session("validAdmin") Then %>
	                    <td align="left"><input type="checkbox" name="coord:NLN:<%= n %>" /></td>
	                <% End If %>
	                <td align="left">
	                <% If repeat = True Then %>
				        R
			        <% End If %>
	                </td>
			        <td align="left"><%= coID %></td>
			        <% If projectValidUser=False Then %>
                        <td align="left"><input type="text" name="coord:date:<%= n %>" value="<%= currentDate %>" readonly /></td>
                    <% ElseIf hortonFlag and Not Session("validAdmin") Then %>
                        <td align="left"><input type="text" name="coord:date:<%= n %>" value="<%= currentDate %>" readonly /></td>
                    <% ElseIf forestarFlag and done="done" Then %>
                        <td align="left"><input type="text" name="coord:date:<%= n %>" value="<%= currentDate %>" readonly /></td>
                    <% Else %>
                        <td align="left"><input class="datepicker" type="text" name="coord:date:<%= n %>" value="<%= currentDate %>"/></td>
                    <% End If %>
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
                    <% ElseIf rsComm("comment") = "This item was marked done" or _
                        rsComm("comment") = "This item was marked NLN" or _
                        rsComm("comment") = "This item was marked complete" or _
                        rsComm("comment") = "This item was marked incomplete" Then %>
	                    <td></td>
	                <% Else %>
	                    <td><button type="button"><a href="viewOpenItemComments.asp?coID=<%=coID%>" target="_blank">V</a></button></td>
	                <% End If %>
                    </tr>
                    <tr><td colspan="10"></td></tr>
                    <% n = n + 1
                    End If
			        rsCoord.MoveNext 
	 	        LOOP 'loop coordinates 
            End If
            If start_n <> n Then %>
                <tr><td colspan="10"><hr /></td></tr>
            <% End If
            RS0.MoveNext
         LOOP 'loop inpection reports
    End If%>
    </table>
    <center>
    <% If coordCnt = 0 Then %>
        <h3>There are no open items at this time</h3>
    <% End If %>
    <input type="submit" value="submit" /><br /><br />
    <% If siteMapInspecID > 0 Then
        SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& siteMapInspecID
        'Response.Write(SQL3)
        SET RS3=connSWPPP.execute(SQL3)
        IF NOT(RS3.EOF) THEN 
            sitemap_link = "http://www.swppp.com/images/sitemap/"& TRIM(RS3("oImageFileName"))%>
	        <a href="<%=sitemap_link%>">link for site map</a><br />
        <% END IF
        IF projectValidUser and (hortonFlag or forestarFlag) THEN
            inspections_link = "http://swppp.com/views/inspections.asp?projID=" & projectID & "&projName=" & projectName & "&projPhase=" & projectPhase %>
            <a href="<%=inspections_link%>">sign off on reports</a><br/>
        <% END IF
    END IF 
    if show_debug then
        Response.Write(dbgBody)
    end if %>
    </center>
</form>
<br /><br />
</body>
</html>

<% connSWPPP.Close
SET connSWPPP=nothing %>
	