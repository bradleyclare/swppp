<%
If Not Session("validAdmin") and not Session("validInspector") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("maintain/loginUser.asp")
End If
%>

<html>
<head>
    <title>SWPPP INSPECTIONS :: Admin</title>
    <link rel="stylesheet" type="text/css" href="../global.css">
</head>

<body>

    <!-- #include file="adminHeader.inc" -->
	
        <h1>Welcome <%= Session("firstName") %> <%= Session("lastName") %></h1>
        <h3>You have valid rights as a
        <% If Session("validAdmin") then %>&nbsp;Admin<% end if %>
        <% If Session("validDirector") then %>&nbsp;Director<% end if %>
        <% If Session("validInspector") then %>&nbsp;Inspector<% end if %>.</h3>
		<p><% Request.ServerVariables("path_info") %></p>
    </div>
</body>
</html>
