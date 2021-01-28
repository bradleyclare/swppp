<%@ Language="VBScript" %><%

If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If

recordOrd = Request("orderBy")
If recordOrd = "" Then
	recordOrd = "lastName"
End If

%><!-- #include file="../connSWPPP.asp" --><%

If Request.Form.Count > 0 Then
    if Request.Form("activate_users") = "activate users" then
        for n = 0 to 999 step 1
            if Trim(Request("user:userID:" & CStr(n))) = "" then
                exit for
            end if
            'Response.Write(CStr(n) & " d-" & Request("user:activate:"& CStr(n)) & " id-" & Request("user::userID:"& CStr(n)) & "</br>")
            if Request("user:activate:"& CStr(n)) = "on" then
                userID = Request("user:userID:" & CStr(n))
                sqlupdate = "UPDATE Users SET active=1 WHERE userID=" & userID
                'response.Write(sqlupdate)
                connSWPPP.execute(sqlupdate)
            End If
        next
    End If	
End If 

SQLSELECT = "SELECT * FROM Users WHERE active=0 ORDER BY " & recordOrd
'Response.Write(SQLSELECT & "<br>")
Set connUsers = connSWPPP.Execute(SQLSELECT)

recCount = 0
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Activate Users for Admins</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>">
<table width="100%" border="0">
	<tr><td><br><h1>Activate Users</h1></td></tr></table>
    <input type="submit" name="activate_users" value="activate users" /></br></br>
<table width="100%" border="0">
	<tr><th><b>Count</b></th>
		<th><b>First Name</b></a></th>
		<th><b>Last Name</b></a></th>
		<th><b>Rights</b></a></th>
        <th><b>Activate</b></th></tr>
<%
	If connUsers.EOF Then
		Response.Write("<tr><td colspan='5' align='center'><b><i>There " & _
			"are currently no inactive users.</i></b></td></tr>")
	Else
		altColors="#ffffff"
		n = 0
		Do While Not connUsers.EOF
			recCount = recCount + 1 %>
            <input type="hidden" name="user:userID:<%= n %>" value="<%= connUsers("userID") %>" />
	        <tr align="center" bgcolor="<%= altColors %>"> 
		    <td><%= recCount %></td>
		    <td><%= Trim(connUsers("firstName")) %></td>
		    <td><%= Trim(connUsers("lastName")) %></a></td>
		    <td><%= connUsers("rights") %></td>
            <td><input type="checkbox" name="user:activate:<%= n %>" /></td></tr>
            <%	If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			n = n + 1
            connUsers.MoveNext
		Loop
	End If ' END No Results Found

connUsers.Close
Set connUsers = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
</table>
</form>
</body>
</html>
