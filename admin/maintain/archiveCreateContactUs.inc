<%
'--	We are going to create Report.html as a textFile one line at a time -----------------
'--	We will do this by calling a function in this include file that will create the file -----
function createContactUs(repDest)

DIM repFSO, repFile, repFolder, fc, f1, xArr, folderName

SET repFSO=CreateObject("Scripting.FileSystemObject") 
SET repFile= repFSO.OpenTextFile(repDest &"ContactUs.html",2,True) 
SET repFolder= repFSO.GetFolder(repDest)

repFile.WriteLine("<html><head><title>SWPPP INSPECTIONS :: Contact Us</title><link rel='stylesheet' href='global.css' type='text/css'></head>")
repFile.WriteLine("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'></head>")
repFile.WriteLine("<body bgcolor='#FFFFFF' text='#000000'>")
repFile.WriteLine("<table width='90%' border=0 cellpadding=0 cellspacing=0 align=center bgcolor='#FFFFFF'>")
repFile.WriteLine("	<tr><td class='header' valign=bottom align=center rowspan=2 width=120>")
repFile.WriteLine("		<img src='images/swpppweb2.jpg' border=0 vspace=5 hspace=5 alt='SWPPP Inspections'></td>")
repFile.WriteLine("		<td nowrap valign=bottom bgcolor='#006699'><br><br><br><br></td>")
repFile.WriteLine("	</tr><tr bgcolor='#999999'>")
repFile.WriteLine("		<td><div class='white'><a href='Reports.html' class='head'>Home</a> &nbsp; | &nbsp; ")
repFile.WriteLine("			<a href='AboutUs.html' class='head'>About Us</a> &nbsp; | &nbsp; ")
repFile.WriteLine("			<a href='ContactUs.html' class='head'>Contact Us</a> &nbsp; | &nbsp; ")
repFile.WriteLine("	</tr></TABLE><table width='90%' border=0 cellpadding=0 cellspacing=0 align=center bgcolor='#FFFFFF'>")
repFile.WriteLine("	<tr><td style='margin-left:10%;' colspan=2><br><br>")
DIM objFSO, objFile
	base_path= Server.Mappath("../../static")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(base_path & "/contact.txt") 
repFile.WriteLine(objFile.ReadAll) 
	Set objFSO = Nothing
	objFile.Close
	Set objFile = Nothing
repFile.WriteLine("</div></td></tr></table></body></html>")
repFile.Close
END function %>