<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

groupID = Request("ID")
del     = Request("del") 

%> <!-- #include file="../connSWPPP.asp" --> <%

If del = 1 Then
    SQLDELETE = "DELETE FROM Groups WHERE groupID =" & groupID
    'Response.Write(SQLDELETE & "<br>")
    connSWPPP.Execute(SQLDELETE)
End If

If Request.Form.Count > 0 Then	
    Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	

    'add new group
    newGroupName= strQuoteReplace(Request("newGroupName"))

    SQLINSERT = "INSERT INTO Groups (groupName) VALUES ('" & newGroupName & "')"
    'Response.Write(SQLINSERT & "<br>")
    connSWPPP.Execute(SQLINSERT)
End If

SQLSELECT = "SELECT groupID, groupName FROM Groups ORDER BY groupName"
'Response.Write(SQLSELECT & "<br>")
Set connGroups = connSWPPP.Execute(SQLSELECT)
recCount = 0
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Manage Groups</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->

<h1>Manage Groups</h1>
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>" onsubmit="return isReady(this)";>
    <input type="text" value="" name="newGroupName" />
    <input type="submit" value="Add New Group" name="submit" />
</form>
<table width="50%" border="0">
	<tr><th width="10%"><b>GroupID</b></th>
		<th width="80%"><b>GroupName</b></th>
		<th width="10%"><b>Delete</b></th></tr>
    <% If connGroups.EOF Then
		Response.Write("<tr><td colspan='3' align='center'><b><i>There are currently no groups.</i></b></td></tr>")
	Else
		altColors="#ffffff"
		
		Do While Not connGroups.EOF
			recCount = recCount + 1 %>
	        <tr align="center" bgcolor="<%= altColors %>"> 
		        <td><%= connGroups("groupID") %></td>
		        <td><%= Trim(connGroups("groupName")) %></td>
		        <td><a href="manageGroups.asp?ID=<%= connGroups("groupID") %>&del=1">delete</a></td></tr>
            <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			connGroups.MoveNext
		Loop
	End If ' END No Results Found
connGroups.Close
Set connGroups = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
</table>
</body>
</html>
