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

<html>
    <head>
	    <title>SWPPP INSPECTIONS</title>
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	    <link rel="stylesheet" type="text/css" href="global.css">
    </head>
    <body>
        <!-- #include file="header3.inc" -->

        <%= objFile.ReadAll %>
    </body>
</html>

<%
objFile.Close
%>