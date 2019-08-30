<%
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
' Response.Write(base_path & "/contact.txt")
Set objFile = objFSO.OpenTextFile(base_path & "/contact.txt")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>SWPPP INSPECTIONS : Contact Us</title>
<link rel="stylesheet" href="../global.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="../header2.inc" -->
<% = objFile.ReadAll %>
<br/>
<br/>
<span id="siteseal"><script async type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=WAZqkjKfwncrXQiy57BnnkkIp0xnpa50j7Om4owvXUaaZQu6tQU4wBV9R1iL"></script></span>
</body></html>
<%
Set objFSO = Nothing

objFile.Close
Set objFile = Nothing
%>