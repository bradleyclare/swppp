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
Set objFile = objFSO.OpenTextFile(base_path & "/about.txt")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
    <head>
        <title>SWPPP INSPECTIONS :: About Us</title>
        <link rel="stylesheet" type="text/css" href="global.css">
    </head>
    <body>
        <!-- #include file="header3.inc" -->

        <% = objFile.ReadAll %>
    </body>
</html>

<%
objFile.Close
%>