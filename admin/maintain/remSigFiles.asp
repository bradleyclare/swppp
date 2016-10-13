<%@  language="VBScript" %>
<%


folderName = "signatures"
displaySize = " : Signatures"

base_path = server.mappath("/")

If Request.Form.Count > 0 Then
	sRemCount = Request("sRemCount")
	
	baseDir = base_path & "\images\" & folderName & "\"
    Response.Write(baseDir)
	Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
	
	For i = 1 To sRemCount
		remSigFile = Request("remSigFile" & i)
		
		' Response.Write(remSigFile & "<br>")
		If remSigFile <> "" Then folderSvrObj.DeleteFile(baseDir & remSigFile) End If
		
	Next
	
	Set folderSvrObj = Nothing
	
	Response.Redirect(Request.ServerVariables("path_info"))
	
End If
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
    <title>SWPPP INSPECTIONS<% = displaySize %> : Removal</title>
    <link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<table width="100%" border="0">
    <tr>
        <td>
            <h1>Remove<% = displaySize %></h1>
        </td>
    </tr>
</table>
<%

baseDir = base_path & "\images\" & folderName & "\"
Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")

Set objImageDir = folderSvrObj.GetFolder(baseDir)
Set remSigFiles = objImageDir.Files
%>
<table width="100%" border="0">
    <tr>
        <td align="center">Please select signature files not associated in database 
			for removal.</td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
</table>
<table width="100%" border="0">
    <form method="post" action="<% = Request.ServerVariables("script_name") %>">
        <tr>
            <th width="20%"><b>Signature</b></th>
            <th width="35%"><b>File Name</b></th>
            <th width="35%"><b>File Date</b></th>
            <th width="10%"><b>Select</b></th>
        </tr>
        <tr>
            <td colspan="4">
                <img src="../../images/dot.gif" width="5" height="5"></td>
        </tr>
        <!-- #include virtual="admin/connSWPPP.asp" -->
        <%
sRemCount = 0

altColors = "#e5e6e8"

For Each Item In remSigFiles
	sFileName = Item.Name
	sFileDate = Item.DateCreated
	
	sigSQLSELECT = "SELECT userID" & _
		" FROM Users " & _
		" WHERE signature = " & "'" & sFileName & "'"
		
	' Response.Write("sigSQLSELECT")
	Set rsSignature = connSWPPP.Execute(sigSQLSELECT)
	
	If rsSignature.EOF Then
		
		sRemCount = sRemCount + 1
        %>
        <tr bgcolor="<% = altColors %>" align="center" nowrap>
            <td height="90">
                <img src="../../images/<% = folderName & "/" & sFileName %>" width="80" height="60" alt="<% = sFileName %>"></td>
            <td>
                <% = sFileName %>
            </td>
            <td>
                <% = sFileDate %>
            </td>
            <td>
                <input type="checkbox" name="remSigFile<% = sRemCount %>" value="<% = sFileName %>"></td>
        </tr>
        <%
		If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
		
		rsSignature.Close
		Set rsSignature = Nothing
		
	End If
	
Next

connSWPPP.Close
Set connSWPPP = Nothing

Set folderSvrObj = Nothing
Set objImageDir = Nothing
Set remSigFiles = Nothing

If sRemCount <> 0 Then
        %>
        <tr align="center">
            <td colspan="4">&nbsp;<br>
                <input type="submit" value="Remove"><br>
                &nbsp;</td>
        </tr>
        <%
Else
	Response.Write("<tr><td colspan='4' align='center'>&nbsp;<br><b><i>Sorry no " & _
		"current unassociated signatures at this time.</i></b><br>&nbsp;</td></tr>")
End If
        %>
        <input type="hidden" name="sRemCount" value="<% = sRemCount %>">
    </form>
</table>
</body>
</html>
