<%
' Get text from a file and inserts into page
' file is editable using 

If Session("email")="" then 
	logUser=Request.ServerVariables("REMOTE_ADDR")
else
	logUser=Session("email")
end if

SQLINSERT =	"INSERT INTO Log (pageName, dateEntered, email" &_
			") VALUES (" &_
			"'About Us'" &_
			", '" & FormatDateTime(now,vbLongDateTime) & "'" &_
			", '" & logUser & "'" &_
			")"
'Response.Write(SQLINSERT & "<br>")
%>	<!-- #INCLUDE FILE="../admin/connSW.asp" --> <%
connSS.execute(SQLINSERT)
connSS.close 

base_path = server.mappath(".")

' iomode settings
ForReading = 1
ForWriting = 2
ForAppending = 8

'format settings
TristateUseDefault = -2
TristateTrue = -1
TristateFalse = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
'response.write(base_path & "/about.txt")
Set objFile = objFSO.OpenTextFile(base_path & "/about.txt")
%>

<html>
<head>
<title>SWPPP INSPECTIONS : About Us</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="global.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!--#include file="header.inc" -->
<tr bgcolor="#FFFFFF"><td colspan="2"><br><br>
	<%= objFile.ReadAll %>
<br><br><br></td></tr>
<tr><td>
<span id="siteseal"><script async type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=WAZqkjKfwncrXQiy57BnnkkIp0xnpa50j7Om4owvXUaaZQu6tQU4wBV9R1iL"></script></span>
</td></tr>
</table>	  
</body>
</html>
