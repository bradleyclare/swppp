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
'response.write(base_path & "/about.txt")
Set objFile = objFSO.OpenTextFile(base_path & "/static/about.txt")
%>

<html>
<head>
	<title>SWPPP INSPECTIONS</title>
	<link rel="stylesheet" type="text/css" href="global.css">
    <link rel="stylesheet" type="text/css" href="./css/bootstrap.min.css" />
	<link rel="stylesheet" type="text/css" href="./css/carousel.css" />
	<link rel="stylesheet" type="text/css" href="./css/my_bootstrap.css" />
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
    <title>SWPPP INSPECTIONS : Home Page</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="../global.css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<!-- #include file="header3.inc" -->
<table><tr><td calspan=*>
<SCRIPT>
<!--
if ((navigator.appVersion.indexOf("MSIE") > 0)
  && (parseInt(navigator.appVersion) >= 4)) {
    var sText = "<font size=-2><SPAN STYLE='color:blue;cursor:hand;' onclick='window.external.AddFavorite(location.href,document.title);'>Add this page to your favorites</SPAN></font>";
    document.write(sText);
}
//-->
</SCRIPT>
</td></tr>
<tr><td calspan=*>
<%= objFile.ReadAll %>
</td></tr>
</table>
</body>
</html>