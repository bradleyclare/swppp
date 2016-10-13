<%
'--	We are going to create each Default Page as a textFile one line at a time -----------------
'--	We will do this by calling a function in this include file that will create the file -----
function createDefault(inspecID, repDest)

DIM repFSO, repFile, repFolder, fc, f1

SET repFSO=CreateObject("Scripting.FileSystemObject") 
SET repFile= repFSO.OpenTextFile(repDest &"Default.html",2,True) 
SET repFolder= repFSO.GetFolder(repDest)

SQLa = "SELECT inspecDate, projectName, projectPhase FROM Inspections  WHERE inspecID = " & inspecID	
Set RSa = connSWPPP.Execute(SQLa)
SQLb="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=-1 ORDER BY oitSortByVal asc"
SET RSb=connSWPPP.execute(SQL1)
inspecDate = Trim(RSa("inspecDate"))
projectName = Trim(RSa("projectName") &" "& RSa("projectPhase"))
RSa.close
SET RSa=nothing

repFile.WriteLine("<html><head><title>SWPPP INSPECTIONS : "& projectName &" on "& inspecDate &"</title>")
repFile.WriteLine("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'><link rel='stylesheet' href='..\global.css' type='text/css'></head>")
repFile.WriteLine("<body bgcolor='#FFFFFF' text='#000000'>")
repFile.WriteLine("<table width='90%' border=0 cellpadding=0 cellspacing=0 align=center bgcolor='#FFFFFF'>")
repFile.WriteLine("	<tr><td class='header' valign=bottom align=center rowspan=2 width='120'>")
repFile.WriteLine("		<img src='../images/swpppweb2.jpg' border=0 vspace=5 hspace=5 alt='SWPPP Inspections'></td>")
repFile.WriteLine("		<td nowrap valign=bottom bgcolor='#006699'><br><br><br><br></td>")
repFile.WriteLine("	</tr><tr bgcolor='#999999'>")
repFile.WriteLine("		<td><div class='white'><a href='..\Reports.html' class='head'>Home</a> &nbsp; | &nbsp; ")
repFile.WriteLine("			<a href='..\AboutUs.html' class='head'>About Us</a> &nbsp; | &nbsp; ")
repFile.WriteLine("			<a href='..\ContactUs.html' class='head'>Contact Us</a> &nbsp;")
repFile.WriteLine("	</tr></TABLE><table width='90%' border=0 cellpadding=0 cellspacing=0 align=center bgcolor='#FFFFFF'>")
repFile.WriteLine("	<tr><td style='margin-left:10%;' colspan=2><br><br>")
repFile.WriteLine("<h1>Inspection for "& projectName &" on "& inspecDate &"</h1>")
repFile.WriteLine("<br><p class='indent30'><a href='FullReport.html' target='_blank'>Report</a><br>")

SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=-1 ORDER BY oitSortByVal asc"
SET RS1=connSWPPP.execute(SQL1)
DO WHILE NOT RS1.EOF 
	SQLa="sp_oImagesByType "& inspecID &",'"& RS1("oitID") &"'" 
	SET RSa=connSWPPP.Execute(SQLa)
	cnt1=1
	curOITDesc=""
	IF NOT(RSa.EOF) THEN repFile.WriteLine("<p class='indent30'>") END IF
 	DO WHILE NOT(RSa.EOF) 
		fileDesc= TRIM(RSa("oitDesc"))
		thisFileDesc=fileDesc
		if curOITDesc=fileDesc then 
			cnt1=cnt1+1
		else
			cnt1=1
			curOITDesc=fileDesc
		end if
		fileName= Trim(RSa("oImageFileName"))
		fileExt= "."& RIGHT(fileName,(LEN(fileName)-InStr(fileName,".")))
		if (cnt1>1) then thisFileDesc=thisFileDesc &" "& cnt1 end if 
		repFile.WriteLine("<a href='"& thisFileDesc & fileExt &"' target='_blank'>"& thisFileDesc &"</a><br>")
 		RSa.MoveNext 
	LOOP 
	RS1.MoveNext
LOOP

imgSQLSELECT = "SELECT imageID, largeImage, smallImage, description FROM Images WHERE inspecID = " & inspecID	
Set rsImages = connSWPPP.execute(imgSQLSELECT)
If Not rsImages.EOF Then
	repFile.WriteLine("<div class='indent30'><b>Site Images:</b><br><br>")
	repFile.WriteLine("<table cellspacing=0 cellpadding=4 width='90%' border=0>")
	repFile.WriteLine("	<tr>")
	Do While Not rsImages.EOF 
		iDataRows = iDataRows + 1
		If iDataRows > 3 Then
			repFile.WriteLine("</tr><tr>")
			iDataRows = 1
		End If
		repFile.WriteLine("	<td align=center><a href='lg/"& Trim(rsImages("largeImage")) &"' target='_blank'>"& Trim(rsImages("description")) &"<br>")
			If Right(Trim(rsImages("smallImage")),3)="pdf" then 
				repFile.WriteLine("<img src='../images/acrobat.gif' width=87 height=30 border=0 alt='Acrobat PDF Doc'>")
			else 
				repFile.WriteLine("<img src='"& "sm/" & Trim(rsImages("smallImage")) &"' border=0 alt='"& Trim(rsImages("smallImage")) &"'>")
			end if 
			repFile.WriteLine("</a></td>")
		rsImages.MoveNext
	Loop
End If
repFile.WriteLine("	</tr></table><br><br></div>")
repFile.WriteLine("</td></tr></table></body><br><br><br></html>")
repFile.Close
END function %>