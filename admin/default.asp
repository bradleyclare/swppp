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

<body bgcolor="#FFFFFF">

<!-- #include file="adminHeader.inc" -->
<div class="indent30"><h1>Welcome <%= Session("firstName") %> <%= Session("lastName") %></h1>
You have valid rights as a
<% If Session("validAdmin") then %>&nbsp;Admin<% end if %>
<% If Session("validDirector") then %>&nbsp;Director<% end if %>
<% If Session("validInspector") then %>&nbsp;Inspector<% end if %>.<br><br>


</div>
</td></tr></table>
</BODY>
</HTML>