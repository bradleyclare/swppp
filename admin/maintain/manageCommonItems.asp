<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validDirector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If

itemID = Request("ID")
del     = Request("del") 

%> <!-- #include file="../connSWPPP.asp" --> <%

If del = 1 Then
    SQLDELETE = "DELETE FROM CommonItems WHERE itemID =" & itemID
    'Response.Write(SQLDELETE & "<br>")
    connSWPPP.Execute(SQLDELETE)
End If

If Request.Form.Count > 0 Then	
    Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	

    'add new group
    newItemName= strQuoteReplace(Request("newItemName"))

    SQLINSERT = "INSERT INTO CommonItems (itemName) VALUES ('" & newItemName & "')"
    'Response.Write(SQLINSERT & "<br>")
    connSWPPP.Execute(SQLINSERT)
End If

SQLSELECT = "SELECT itemID, itemName FROM CommonItems ORDER BY itemName"
'Response.Write(SQLSELECT & "<br>")
Set connItems = connSWPPP.Execute(SQLSELECT)
recCount = 0
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Manage Common Items</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->

<h1>Manage Common Items</h1>
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>" onsubmit="return isReady(this)";>
    <input type="text" value="" size="100" name="newItemName" />
    <input type="submit" value="Add New Item" name="submit" />
</form>
<table width="80%" border="0">
	<tr><th width="10%"><b>ItemID</b></th>
		<th width="80%" align="left"><b>ItemName</b></th>
		<th width="10%"><b>Delete</b></th></tr>
    <% If connItems.EOF Then
		Response.Write("<tr><td colspan='3' align='center'><b><i>There are currently no common items.</i></b></td></tr>")
	Else
		altColors="#ffffff"
		
		Do While Not connItems.EOF
			recCount = recCount + 1 %>
	        <tr align="center" bgcolor="<%= altColors %>"> 
		        <td><%= connItems("itemID") %></td>
		        <td align="left"><%= Trim(connItems("itemName")) %></td>
		        <td><a href="manageCommonItems.asp?ID=<%= connItems("itemID") %>&del=1">delete</a></td></tr>
            <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
			connItems.MoveNext
		Loop
	End If ' END No Results Found
connItems.Close
Set connItems = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>
</table>
</body>
</html>
