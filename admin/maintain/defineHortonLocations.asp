<%@ Language="VBScript" %>
<% If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID = Request("inspecID")
outfallFlag = Request("outfall")
IF outfallFlag="" THEN 
	outfallFlag=0
END IF

%><!-- #include file="../connSWPPP.asp" --><%
If Request.Form.Count > 0 Then
    Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	
    for n = 1 to 999 step 1
        if Trim(Request("locationID:" & CStr(n))) = "" then
		    exit for
		end if
        locationID = Request("locationID:"& Cstr(n))
        locationName = TRIM(strQuoteReplace(Request("locationName:"& Cstr(n))))
        If locationName <> "" Then
            If locationID = 0 Then
                insertSQL = "INSERT INTO HortonLocations (inspecID, locationName, isOutfall, answer) VALUES ("& inspecID &", '"& locationName &"', "& outfallFlag &", 1)"
                Response.Write(insertSQL &"</br>")
                connSWPPP.Execute(insertSQL)
            Else
                upateSQL = "UPDATE HortonLocations SET locationName='"& locationName &"', isOutfall="& outfallFlag &" WHERE locationID="& locationID
                Response.Write(upateSQL &"</br>")
                connSWPPP.Execute(upateSQL)
            End If
        End If
    next
End If

SQL1="SELECT * FROM HortonLocations WHERE inspecID="& inspecID &" AND isOutfall="& outfallFlag
'response.Write(SQL1)
Set RS1=connSWPPP.execute(SQL1) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Define Horton Locations</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<% If outfallFlag Then %>
<h1>define outfall locations</h1>
<% Else %>
<h1>define pond locations</h1>
<% End If %>
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?inspecID=<%=inspecID%>&outfall=<%=outfallFlag%>" onsubmit="return isReady(this)";>
<table style="margin-left:auto; margin-right:auto;">
    <tr><th width="5%">#</th>
    <th width="10%">locations</th></tr>
    <% n = 1
	Do While Not RS1.EOF
        locationID = RS1("locationID")
        locationName = RS1("locationName") %>
        <input type="hidden" name="locationID:<%= n %>" value="<%= locationID %>" />
        <tr><td align="center"><%=n%></td>
        <td><input type="text" size="10" name="locationName:<%= n %>" value="<%=locationName %>" /></td></tr>
        <% RS1.MoveNext
        n = n + 1
	Loop

    for m = n to n+9 step 1 %>
        <input type="hidden" name="locationID:<%= m %>" value="0" />
	    <tr><td align="center"><%=m%></td>
        <td><input type="text" size="10" name="locationName:<%= m %>" value="" /></td></tr>
    <% next %>
    <tr><td colspan="2" align="center"><input name="submit_location" type="submit" style="font-size: 20px;" value="submit"/></td></tr>
    <tr></tr>
    <tr><td colspan="2" align="center"><a style="font-size: 15px;" href="hortonQuestions.asp?inspecID=<%=inspecID%>">return to questions page</a></td></tr>
</table>
</form>
</body>
</html>
<% connSWPPP.Close
Set connSWPPP = Nothing %>
