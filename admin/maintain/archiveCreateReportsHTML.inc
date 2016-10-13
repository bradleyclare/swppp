<%
'--	We are going to create Report.html as a textFile one line at a time -----------------
'--	We will do this by calling a function in this include file that will create the file -----
function createReportsHTML(repDest, projectName)

DIM repFSO, repFile, repFolder, fc, f1, f2, xArr, folderName

SET repFSO=CreateObject("Scripting.FileSystemObject") 
SET repFile= repFSO.OpenTextFile(repDest &"Reports.html",2,True) 
SET repFolder= repFSO.GetFolder(repDest)

repFile.WriteLine("<html><head><title>SWPPP INSPECTIONS : Report Dates</title>")
repFile.WriteLine("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'></head>")
repFile.WriteLine("<body bgcolor='#FFFFFF' text='#000000'>")

repFile.WriteLine("<html><head><title>SWPPP INSPECTIONS : Report Dates</title>")
repFile.WriteLine("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>")
repFile.WriteLine("<link href='global.css' rel='stylesheet' type='text/css'></head>")
repFile.WriteLine("<body bgcolor='#FFFFFF' text='#000000'>")
repFile.WriteLine("<table width='90%' border=0 cellpadding=0 cellspacing=0 align=center bgcolor='#FFFFFF'>")
repFile.WriteLine("	<tr><td class='header' valign=bottom align=center rowspan=2 width='120'>")
repFile.WriteLine("		<img src='images/swpppweb2.jpg' border=0 vspace=5 hspace=5 alt='SWPPP Inspections'></td>")
repFile.WriteLine("		<td nowrap valign=bottom bgcolor='#006699'><br><br><br><br></td>")
repFile.WriteLine("	</tr><tr bgcolor='#999999'>")
repFile.WriteLine("		<td><div class='white'><a href='Reports.html' class='head'>Home</a> &nbsp; | &nbsp; ")
repFile.WriteLine("			<a href='AboutUs.html' class='head'>About Us</a> &nbsp; | &nbsp; ")
repFile.WriteLine("			<a href='ContactUs.html' class='head'>Contact Us</a> &nbsp; | &nbsp; ")
repFile.WriteLine("	</tr></TABLE><table width='90%' border=0 cellpadding=0 cellspacing=0 align=center bgcolor='#FFFFFF'>")
repFile.WriteLine("	<tr><td style='margin-left:10%;' colspan=2><br><br>")
repFile.WriteLine("<h1>"& projectName &"</h1></td></tr>")
repFile.WriteLine("<tr><td colspan=2><div style='margin-left:10%;'>")
Set fc = repFolder.SubFolders
tempStr=""
For Each f1 in fc
	IF f1.name<>"Images" THEN
		IF (repFSO.folderExists(repDest&"\"& f1.name &"\sm") AND repFSO.folderExists(repDest&"\"& f1.name &"\lg")) THEN imgStr="&nbsp;<img src='Images\smallcamera.gif'>" ELSE imgStr="" END IF
		folderName= DATEADD("d",#1/1/2000#,f1.name)
		tempStr= "<a href='"& f1.Name &"\Default.html'>"& folderName &"</a>"& imgStr &"<br>" & VbCrLf & tempStr
	END IF
Next
repFile.WriteLine(tempStr)
repFile.WriteLine("</div></td></tr></table></body></html>")
repFile.Close
END function %>