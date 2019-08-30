<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

del     = Request("del") 
groupName = "unknown"

if len(Request.QueryString("ID")) > 0 Then
    userGroupID = Trim(Request("ID"))
Else
    userGroupID = 1
End If

%> <!-- #include file="../connSWPPP.asp" --> <%

If del = 1 Then
    SQLDELETE = "DELETE FROM UserGroups WHERE userGroupID =" & userGroupID
    'Response.Write(SQLDELETE & "<br>")
    connSWPPP.Execute(SQLDELETE)
End If

If Request.Form.Count > 0 Then	
    Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	

    'add new group
    newUserGroupName= strQuoteReplace(Request("newUserGroupName"))

    SQLINSERT = "INSERT INTO UserGroups (userGroupName) VALUES ('" & newUserGroupName & "')"
    'Response.Write(SQLINSERT & "<br>")
    connSWPPP.Execute(SQLINSERT)
End If

SQLSELECT = "SELECT userGroupID, userGroupName FROM UserGroups ORDER BY userGroupID"
'Response.Write(SQLSELECT & "<br>")
Set connGroups = connSWPPP.Execute(SQLSELECT)
recCount = 0
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Manage User Groups</title>
	<link rel="Stylesheet" href="../../global.css" type="text/css" />
</head>
<!-- #include file="../adminHeader2.inc" -->

<h1>Manage User Groups</h1>
<table width="90%">
    <tr>
        <td width="40%" valign="top">
            <form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>" onsubmit="return isReady(this)";>
                <input type="text" value="" name="newUserGroupName" />
                <input type="submit" value="Add New Group" name="submit" />
            </form>
            <table>
	            <tr><th width="80%"><b>userGroupName</b></th>
		            <th width="10%"><b>Delete</b></th></tr>
                <% If connGroups.EOF Then
		            Response.Write("<tr><td colspan='3' align='center'><b><i>There are currently no user groups.</i></b></td></tr>")
	            Else
		            altColors="#ffffff"
		
		            Do While Not connGroups.EOF
			            recCount = recCount + 1 %>
	                    <tr bgcolor="<%= altColors %>"> 
		                    <td><a href="manageUserGroups.asp?ID=<%= connGroups("userGroupID") %>"><%= Trim(connGroups("userGroupName")) %></a></td>
		                    <td><a href="manageUserGroups.asp?ID=<%= connGroups("userGroupID") %>&del=1" onclick="return confirm('Are you sure you want to delete this user group?')">delete</a></td></tr>
                        <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			            If userGroupID = Trim(connGroups("userGroupID")) Then userGroupName = Trim(connGroups("userGroupName")) End If
                        connGroups.MoveNext
		            Loop
	            End If ' END No Results Found
                %>
            </table>
        </td>
        <td width="40%" valign="top">
            <% SQL0 = "SELECT DISTINCT userID, firstName, lastName, userGroupID FROM Users" & _
                    " WHERE userGroupID = '" & userGroupID & "'" & _
                    " ORDER BY lastName"
            'Response.Write(SQL0)
            Set RS0 = connSWPPP.Execute(SQL0) 
            
            SQLSELECT = "SELECT userGroupID, userGroupName FROM UserGroups" & _
                " WHERE userGroupID = '" & userGroupID & "'"
            'Response.Write(SQLSELECT & "<br>")
            Set connGroups = connSWPPP.Execute(SQLSELECT) %>
            
			<% if connGroups.EOF Then %>
				<h5>Not a valid Group</h5>
			<% else %>
            <h3>Users in Group [<%=connGroups("userGroupName")%>]</h3>
            <% If RS0.EOF Then %>
                <h5>No users assigned to group.</h5>
            <% Else %>
            <table>
	            <tr><th width="10%"><b>User ID</b></th>
		            <th width="90%"><b>User Name</b></th></tr>
                <% Do While Not RS0.EOF %>
                    <tr>
                        <td><%=RS0("userID") %></td>
                        <td><%=RS0("firstName") %> <%=RS0("lastName") %></td>
                    </tr>
                <% RS0.MoveNext
                Loop 
            End If 
			End If %>
            </table>
        </td>
    </tr>
</table>
</body>
</html>
