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

projID = Trim(Request("pID"))

%><!-- #include file="../admin/connSWPPP.asp" --><%

Server.ScriptTimeout=1500

currentDate = date()

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
        " VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& currentDate & "', "& inspecID & ", "& projID & ")"   
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
                coID = Request("coord:coID:"& CStr(n))
			    SQLc = "UPDATE Coordinates "& _
			    "SET status=1, completeDate='" & Request("coord:date:"& CStr(n))& "' " & _ 
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
                comment = "This item was marked complete"
                'Response.Write(coID & " - " & userID & " - " & currentDate & " - " & comment)
                SQL3="INSERT INTO CoordinatesComments (coID, comment, userID, date, inspecID, projectID)" &_
				" VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& currentDate & "', "& inspecID & ", "& projID & ")"    
                'response.Write(SQL3)
                Set RS3=connSWPPP.execute(SQL3)
		    End If
            If Request("coord:NLN:"& CStr(n)) = "on" Then
                coID = Request("coord:coID:"& CStr(n))
                SQLc = "UPDATE Coordinates "& _
			    "SET NLN=1, completeDate='" & Request("coord:date:"& CStr(n))& "' " & _ 
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
				" VALUES ( "& coID & ", '" & comment & "', " & userID & ", '"& currentDate & "', "& inspecID & ", "& projID & ")"  
                'response.Write(SQL3)
                Set RS3=connSWPPP.execute(SQL3)
            End If 
	    next	
    'End If
End If

SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projID
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
SQL0 = "SELECT inspecID, inspecDate, reportType," & _
    " projectID, projectName, projectPhase, released, includeItems, compliance, totalItems, completedItems" & _
    " FROM Inspections" & _
    " WHERE projectID = " & projID &_
    " AND includeItems = 1" &_ 
    " AND released = 1" &_
    " AND compliance = 0" &_
    " AND openItemAlert = 1" 
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)
%>

<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")&"?pID="&projID%>" onsubmit="return isReady(this)";>
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
    <font size="+1"><b>Open Items for<br/> (<%=projID %>) <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
    <br /><br />
    <table><tr>
    <td><input type="button" value="Check all Items" onclick="check_all_items(this)" /></td>
    <td><input type="button" value="Un-Check all Items" onclick="uncheck_all_items(this)" /></td>
    <td><input type="text" name="commonDate" class="datepicker" value="<%= currentDate %>" /></td>
    <td><input type="button" value="Apply Date to All" onclick="apply_date_to_all(this)" /></td>
    </tr></table>
    <br/>
    <a href="completedActionItems.asp?pID=<%=projID%>&inspecID=<%=inspecID%>">see Completed Items</a>
    <br/><br/>
    </center>
    <table cellpadding="2" cellspacing="0" border="0" width="100%">
	    <tr><th width="5%" align="left">Complete</th>
            <% If Session("validAdmin") or Session("validDirector") Then %>
            <th width="5%" align="left">NLN</th>
            <% End If %>
            <th width="5%" align="left">Repeat</th>
            <th width="5%" align="left">ID</th>
            <th width="10%" align="left">Completion Date</th>
            <th width="5%" align="left">Age</th>
            <th width="5%" align="left">Report Date</th>
            <th width="20%" align="left">Location</th>
            <th align="left">Action Item</th>
            <th width="2.5%" align="left">Add Note</th>
            <th width="2.5%" align="left">View Note</th>
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
			 coordSQLSELECT = "SELECT coID, coordinates, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD, NLN, parentID FROM Coordinates" &_
                " WHERE inspecID=" & inspecID &_
                " AND status=0" &_
                " AND infoOnly=0" &_
                " ORDER BY orderby"	
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
			        If infoOnly = True or NLN = True Then
	                 do_nothing = 1 
	              Elseif status = false Then 
	                 commSQLSELECT = "SELECT comment, userID, date" &_
		               " FROM CoordinatesComments WHERE coID=" & coID	
	                 Set rsComm = connSWPPP.execute(commSQLSELECT) %>
			           <input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
	                 <input type="hidden" name="coord:inspecID:<%= n %>" value="<%= inspecID %>" />
			           <tr>
			           <td align="left"><input type="checkbox" name="coord:complete:<%= n %>" /></td>
			           <% If Session("validAdmin") or Session("validDirector") Then %>
	                    <td align="left"><input type="checkbox" name="coord:NLN:<%= n %>" /></td>
	                 <% End If %>
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
    <input type="submit" value="Submit" /><br /><br />
    <% If siteMapInspecID > 0 Then
        SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& siteMapInspecID
        'Response.Write(SQL3)
        SET RS3=connSWPPP.execute(SQL3)
        IF NOT(RS3.EOF) THEN 
            sitemap_link = "http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName"))%>
	        <a href='<%=sitemap_link%>'>link for Site Map</a><br />
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
	