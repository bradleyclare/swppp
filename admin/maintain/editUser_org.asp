<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

userID = Request("userID")
companyID = Request("companyID")
%>
<!-- #include virtual="admin/connSWPPP.asp" -->
<%
If Request.Form.Count > 0 Then
	userSQLUPDATE =	"UPDATE Users SET" & _
		" firstName = '" & Request("firstName") & "'" & _
		", lastName = '" & Request("lastName") & "'" & _
		", email = '" & Request("email") & "'" & _
		", rights = '" & Request("rights") & "'" & _
		", pswrd = '" & Request("pswrd") & "'" & _
		", signature = '" & Request("signature") & "'" & _
		" WHERE userID = " & userID
		
	' Response.Write(userSQLUPDATE)
	connSWPPP.Execute(userSQLUPDATE)
	
	connSWPPP.Close
	Set connSWPPP = Nothing
	
	Response.Redirect("viewUsers.asp")
	
Else
	userSQLSELECT = "SELECT firstName, lastName, email" & _
		", pswrd, dateEntered, signature" & _
		" FROM Users" & _
		" WHERE userID = " & userID
		
	' Response.Write(userSQLSELECT)
	Set connUser = connSWPPP.Execute(userSQLSELECT)
	
	dateEntered = connUser("dateEntered")
	firstName = Trim(connUser("firstName"))
	lastName = Trim(connUser("lastName"))
	email = Trim(connUser("email"))
	pswrd = Trim(connUser("pswrd"))
	' rights = Trim(connUser("rights"))
	signature = Trim(connUser("signature"))
	
	connUser.Close
	Set connUser = Nothing
	
End If
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Edit User</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include virtual="admin/adminHeader2.inc" -->
    
<table width="100%" border="0">
	<form action="<%= Request.ServerVariables("script_name") %>" method="post">
		<input type="hidden" name="userID" value="<% = userID %>">
		<tr><td colspan="2"><h1>Edit User</h1></td></tr>
		<tr align="center"> 
			<td colspan="2"><input type="button" value="Delete User" 
			onClick="location='deleteUser.asp?userID=<% = userID %>'; return false";></td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<tr><td width="35%" align="right">Date Entered:</td>
			<td width="65%"><%= dateEntered %></td></tr>
		<tr><td align="right">First Name:</td>
			<td><input type="text" name="firstName" size="20" maxlength="20" 
			value="<% = firstName %>"></td></tr>
		<tr><td align="right">Last Name:</td>
			<td><input type="text" name="lastName" size="20" maxlength="20" 
			value="<% = lastName %>"></td></tr>
		<tr> 
			<td align="right">Email:</td>
			<td><input type="text" name="email" size="30" maxlength="40" 
			value="<% = email %>"></td>
		</tr>
		<tr> 
			<td align="right">Password:</td>
			<td><input type="password" name="pswrd" size="8" maxlength="8" 
			value="<% = pswrd %>"> &nbsp; <input type="button" value="View" onClick="alert('Password: ' + form.pswrd.value)";></td>
		</tr>
		<tr> 
			<td align="right">Rights:</td>
			<td><select name="rights">
					<option value="user"<% If rights = "user" Then %> selected<% End If %>>User</option>
					<option value="director"<% If rights = "director" Then %> selected<% End If %>>Director</option>
					<option value="inspector"<% If rights = "inspector" Then %> selected<% End If %>>Inspector</option>
					<option value="admin"<% If rights = "admin" Then %> selected<% End If %>>Admin</option>
				</select></td>
		</tr>
		<tr> 
			<td align="right">Signature File:</td>
			<td><select name="signature">
<%
' get gif directory
Set folderServerObj = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = folderServerObj.GetFolder("d:\vol\swpppinspections.com\www\htdocs\images\signatures\")
Set gifDirectory = objFolder.Files

For Each gifFile In gifDirectory
	shortenedName = gifFile.Name
%>
					<option value="<% = shortenedName %>" 
				<% If signature = shortenedName Then %>selected<% End If %>>
					<% = shortenedName %>
					</option>
<%
Next

Set objFolder = Nothing
Set gifDirectory = Nothing
%>
				</select>
				&nbsp;&nbsp;
				<input type="button" value="Upload Signature File" 
				onClick="location='upSigEditUser.asp?userID=<% = Request("id") %>'; return false";></td>
		</tr>

<% 'admins have all rights and don't need to select companies
	If rights<>"admin" Then %>
		<tr><td align="right" valign="top"><br>Rights to Company:</td>
			<td><br><table>
				<% 	' select all the companies
					SQLSELECT = "SELECT companyID, companyName" &_
						" FROM Companies ORDER BY companyName"
					SET connCompanies=connSWPPP.execute(SQLSELECT)
					Do While not connCompanies.eof %>
					<tr><td><%= Trim(connCompanies("companyName")) %></td>
						<td><input type="checkbox" name="<%= connCompanies("companyID") %>"
						<% 	'is this user and company associated?
							SQLSELECT = "SELECT * FROM CompanyUsers" &_
								" WHERE userID=" & userID &_
								" AND companyID=" & connCompanies("companyID")
							'Response.Write(SQLSELECT & "<br>")
						 	SET connAssoc=connSWPPP.execute(SQLSELECT)
							If not connAssoc.eof then %>checked<% end if  %>></td></tr>
						<% 	connAssoc.close
						connCompanies.movenext
					loop 
					connCompanies.close
					connSWPPP.close %></table>
					</td></tr>
<% 	end if 'rights<>"admin" %>

		<tr align="center"><td colspan="2"><br><br><input type="submit" value="Edit User"></td></tr>
	</form>
</table>
</body>
</html>
