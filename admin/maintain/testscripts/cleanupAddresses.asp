<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../../loginUser.asp")
End If 
inspecID = Session("inpecID")
IF Request("inspecID")<>"" THEN 
	inspecID = Request("inspecID") 
	Session("inspecID")=inspecID
END IF %>
<!-- #include file="../../connSWPPP.asp" -->
<% If Request.Form.Count > 0 Then
	If Request("clean-up") = "Clean Up Records" Then
		totalItems = 0
		for n = 1 to 9999 step 1
			if Trim(Request("coord:" & CStr(n))) = "" then
		        exit for
		    end if
			totalItems = totalItems + 1
			id = Trim(Request("id:" & CStr(n)))
			coordinate = Trim(Request("coord:" & CStr(n)))
			address = Trim(Request("address:" & CStr(n)))
			SQLc = "UPDATE Addresses SET locationName ='" & coordinate & _
			"', address ='" & address & "' WHERE addressID =" & id & ";"
			'Response.Write(SQLc)
			'Response.Write("<br/>")
			connSWPPP.execute(SQLc)
		next
    End If
End If 
%>
<html>
<head>
	<title>SWPPP INSPECTIONS : Clean Up Addresses</title>
	<link rel="stylesheet" type="text/css" href="../../../global.css">
</head>
<body>
<!-- #include file="../../adminHeader3.inc" -->
<h1>Clean Up Addresses</h1>
<% addressSQLSELECT = "SELECT addressID, locationName, address FROM Addresses ORDER BY locationName"
'Response.Write(addressSQLSELECT)
Set rsAddress = connSWPPP.execute(addressSQLSELECT) %>
<form id="theForm" method="post" action="<% = Request.ServerVariables("script_name") %>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
	<table><tr>
	<td><a href="editReport.asp?inspecID=<%=inspecID%>"><button type="button">Return to Report</button></a></td>
	<td><input type="submit" name="clean-up" value="Clean Up Records"/></td>
	</tr></table>
	
	<h3>Proposed Changes to Address Database</h3>
	<% if not rsAddress.EOF Then %>
	<table width="90%">
	<tr><th width="10%">ID</th><th width="22%">Coordinates</th><th width="22%">Mod Coordinates</th><th width="22%">Address</th><th width="22%">Mod Address</th></tr>
	<% n = 0
	Do While Not rsAddress.EOF 
		id               = TRIM(rsAddress("addressID")) 
		name             = TRIM(rsAddress("locationName")) 
		addname          = TRIM(rsAddress("address")) 
		mod_name         = name
		mod_addname      = addname
		name_class       = ""
		addname_class    = ""
		mod_flag         = False
		'look for single digits in the first number and add a leading zero to fix sorting
		parts = Split(name)
		for i=0 to Ubound(parts) step 1
			if isnumeric(parts(i)) Then
				orig = parts(i)
				num = CInt(parts(i))
				if num < 10 and len(orig) < 2 THEN
					mod_flag = True
					name_class = "red"
					newp = "0" & parts(i)
					parts(i) = newp
					'mod_name = Replace(name,orig,newp,1,1)
					'name = mod_name
				End If
			End If
		Next
		mod_name = join(parts)
		'remove any trailing commas and spaces in the address name
		for i=0 to len(mod_addname) step 1
			last_char = right(mod_addname,1)
			if last_char = "," OR last_char = " " THEN
				mod_flag = True
				addname_class = "red"
				mod_addname = left(mod_addname,len(mod_addname)-1)
			End If
		Next
		if mod_flag Then 
			n = n + 1 %>
			<tr><td><input type="text" name="id:<%= n %>" value="<%=id%>" readonly/></td>
			<td><%=name%></td>
			<td><input type="text" class="<%=name_class%>" name="coord:<%= n %>" value="<%=mod_name%>"/></td>
			<td><%=addname%></td>
			<td><input type="text" class="<%=addname_class%>" name="address:<%= n %>" value="<%=mod_addname%>"/></td></tr>
		<% End If
		rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst %>
	</table>
	<% if n = 0 Then %>
		<h5>No Changes Proposed</h5>
	<% End If
	Else %>
		<h5>No Address Info Found</h5>
	<% End If %>
	</div>
</form>
<hr>
<% connSWPPP.Close 
Set connSWPPP = Nothing %>	
</body>
</html>