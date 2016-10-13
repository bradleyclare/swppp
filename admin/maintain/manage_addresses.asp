<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID = Session("inpecID")
IF Request("inspecID")<>"" THEN 
	inspecID = Request("inspecID") 
	Session("inspecID")=inspecID
END IF %>
<!-- #include file="../connSWPPP.asp" -->
<% inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, projectAddr, projectCity, projectState" & _
		", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, i.projectID, compName" & _
		", compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax" & _
		", contactEmail, reportType, inches, bmpsInPlace, sediment, userID, includeItems, compliance, totalItems, completedItems" & _
		" FROM Inspections as i, Projects as p" & _
		" WHERE i.projectID = p.projectID AND inspecID = " & inspecID
'Response.Write(inspecSQLSELECT & "<br>")
Set rsReport = connSWPPP.execute(inspecSQLSELECT)
projectID = rsReport("projectID")
projectName = rsReport("projectName")

Response.Write(request("submit"))
If Request.Form.Count > 0 Then	
	If request("delete") = "Delete All Records" Then	
	    SQLc = "DELETE FROM Addresses WHERE projectID = " & projectID & ";"
		'Response.Write(SQLc)
		connSWPPP.execute(SQLc)
    End If
	If request("save") = "Save Records to Database" Then	
		totalItems = 0
		for n = 0 to 999 step 1
			if Trim(Request("coord:" & CStr(n))) = "" then
		        exit for
		    end if
			totalItems = totalItems + 1
			coordinate = Trim(Request("coord:" & CStr(n)))
			address = Trim(Request("address:" & CStr(n)))
			SQLc = "INSERT INTO Addresses (projectID, locationName, address) " & _
			"Values (" & projectID & ",'" & coordinate & "','" & address & "');"
			'Response.Write(SQLc)
			connSWPPP.execute(SQLc)
		next
    End If
End If 
%>
<html>
<head>
	<title>SWPPP INSPECTIONS : Manage Addresses</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<link href="../../css/jquery-ui.min.css" rel="stylesheet" type="text/css"/>
	<link href="../../css/jquery-ui.structure.min.css" rel="stylesheet" type="text/css"/>
	<link href="../../css/jquery-ui.theme.min.css" rel="stylesheet" type="text/css"/>
	<script src="../../js/jquery.js" type="text/javascript"></script>
	<script src="../../js/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" >
$(document).ready( function () {
	
	var fileInput = document.getElementById("choose_file"),

    readFile = function () {
        var reader = new FileReader();
        reader.onload = function () {
			var csv = reader.result;
			var allTextLines = csv.split(/\r\n|\n/);
			var html = '<table><tr><th>Coordinates</th><th>Address</th></tr>';
			for (var i=0; i<allTextLines.length; i++) {
				html += '<tr>';
				var data = allTextLines[i].split(',');
                for (var j=0; j<data.length; j++) {
					if (j==0){
						name = 'coord:' + i;
					} else {
						name = 'address:' + i;
					}
                    html += '<td><input type="edit" name="' + name + '" value="' + data[j] + '" /></td>';
                }
                html += '</tr>';
			}
			html += '</table>';
            document.getElementById('out').innerHTML = html;
        };
		reader.onerror = function(){ alert('Unable to read ' + fileInput.fileName); };
        // start reading the file. When it is done, calls the onload event defined above.
        reader.readAsText(fileInput.files[0]);
    };

	fileInput.addEventListener('change', readFile);
});
</script>
	
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Manage Addresses for <%=projectName%></h1>
<% addressSQLSELECT = "SELECT addressID, locationName, address FROM Addresses WHERE projectID=" & projectID & " ORDER BY locationName"
'Response.Write(addressSQLSELECT)
Set rsAddress = connSWPPP.execute(addressSQLSELECT) %>
<form id="theForm" method="post" action="<% = Request.ServerVariables("script_name") %>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
	<input type="hidden" name="projectID" value="<%=projectID%>"/>
	
	<table><tr>
	<td><a href="editReport.asp?inspecID=<%=inspecID%>"><button type="button">Return to Report</button></a></td>
	<td><input type="submit" name="delete" value="Delete All Records"/></td>
	<td><input type="file" name="add" id="choose_file" value=""/></td>
	<td><input type="submit" name="save" value="Save Records to Database"/></td>
	<td><a href="clean_up_addresses.asp?inspecID=<%=inspecID%>"><button type="button">Clean Up Addresses</button></a></td>
	</tr></table>
	
	<div class="fl half">
	<h3>Address Information in Database</h3>
	<% if not rsAddress.EOF Then %>
	<table width="90%">
	<tr><th width="10%">ID</th><th width="40%">Coordinates</th><th width="40%">Address</th></tr>
	<% Do While Not rsAddress.EOF 
		id = TRIM(rsAddress("addressID")) 
		name = TRIM(rsAddress("locationName")) 
		addname = TRIM(rsAddress("address")) %>
		<tr><td><%=id%></td><td><%=name%></td><td><%=addname%></td></tr>
		<% rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst %>
	</table>
	<% Else %>
		<h5>No Addresses Found</h5>
	<% End If %>
	</div>
	<div class="fl half">
	<h3>Address Information from File <%=filename%></h3>
	<div id="out">
	</div>
	</div>
	<div class="cleaner"></div>
</form>
<hr>
<% connSWPPP.Close 
Set connSWPPP = Nothing %>	
</body>
</html>