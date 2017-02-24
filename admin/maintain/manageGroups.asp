<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

del     = Request("del") 
groupName = "unknown"

if len(Request.QueryString("ID")) > 0 Then
    groupID = Trim(Request("ID"))
Else
    groupID = 1
End If

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
	<link rel="Stylesheet" href="../../global.css" type="text/css" />
</head>
<!-- #include file="../adminHeader2.inc" -->

<h1>Manage Groups</h1>
<table width="90%">
    <tr>
        <td width="40%" valign="top">
            <form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>" onsubmit="return isReady(this)";>
                <input type="text" value="" name="newGroupName" />
                <input type="submit" value="Add New Group" name="submit" />
            </form>
            <table>
	            <tr><th width="10%"><b>GroupID</b></th>
		            <th width="80%"><b>GroupName</b></th>
		            <th width="10%"><b>Delete</b></th></tr>
                <% If connGroups.EOF Then
		            Response.Write("<tr><td colspan='3' align='center'><b><i>There are currently no groups.</i></b></td></tr>")
	            Else
		            altColors="#ffffff"
		
		            Do While Not connGroups.EOF
			            recCount = recCount + 1 %>
	                    <tr bgcolor="<%= altColors %>"> 
		                    <td><%= connGroups("groupID") %></td>
		                    <td><a href="manageGroups.asp?ID=<%= connGroups("groupID") %>"><%= Trim(connGroups("groupName")) %></a></td>
		                    <td><a href="manageGroups.asp?ID=<%= connGroups("groupID") %>&del=1">delete</a></td></tr>
                        <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			            If groupID = Trim(connGroups("groupID")) Then groupName = Trim(connGroups("groupName")) End If
                        connGroups.MoveNext
		            Loop
	            End If ' END No Results Found
                %>
            </table>
        </td>
        <td width="40%" valign="top">
            <% SQL0 = "SELECT DISTINCT projectID, projectName, projectPhase, groupName FROM Inspections" & _
                    " WHERE groupName = '" & groupName & "'" & _
                    " ORDER BY projectName"
            'Response.Write(SQL0)
            Set RS0 = connSWPPP.Execute(SQL0) %>
            <h3>Projects in Group [<%=groupName%>]</h3>
            <% If RS0.EOF Then %>
                <h5>No projects assigned to group.</h5>
            <% Else %>
            <table>
	            <tr><th width="10%"><b>ProjectID</b></th>
		            <th width="90%"><b>ProjectName</b></th></tr>
                <% Do While Not RS0.EOF %>
                    <tr>
                        <td><%=RS0("projectID") %></td>
                        <td><%=RS0("projectName") %> <%=RS0("projectPhase") %></td>
                    </tr>
                <% RS0.MoveNext
                Loop 
            End If %>
            </table>
        </td>
    </tr>
</table>
</body>
</html>

<%
RS0.Close
Set RS0 = Nothing

connGroups.Close
Set connGroups = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
