<%
'--	We are going to create each full Report as a textFile one line at a time -----------------
'--	We will do this by calling a function in this include file that will create the file -----
function createFullReportHTML(inspecID, repDest)

inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, projectAddr, projectCity, projectState, " & _
	"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " & _
	"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, " & _
	"contactFax, contactEmail, reportType, inches, bmpsInPlace, " & _
	"sediment, narrative, firstName, lastName, signature, qualifications" & _
	" FROM Inspections, Projects, Users WHERE inspecID = " & inspecID & _
	" AND Inspections.projectID = Projects.projectID AND Inspections.userID = Users.userID"
Set rsInspec = connSWPPP.Execute(inspecSQLSELECT)
bmpsInPlace = "No"
If rsInspec("bmpsInPlace") = "1" Then bmpsInPlace = "Yes" End If
sediment = "No"
If rsInspec("sediment") ="1" Then sediment = "Yes" End If
reportType = Trim(rsInspec("reportType"))
inches = rsInspec("inches")
printName = Trim(rsInspec("firstName")) & " " & Trim(rsInspec("lastName"))
narrative= TRIM(rsInspec("narrative"))
IF IsNull(narrative) THEN narrative="" END IF
qualifications= TRIM(rsInspec("qualifications"))
IF IsNull(qualifications) THEN qualifications="" END IF 
SET repFSO=CreateObject("Scripting.FileSystemObject") 
SET repFile= repFSO.OpenTextFile(repDest,2,True) 
repFile.WriteLine("<html><head><title>SWPPP INSPECTIONS - Print Report</title></head>")
repFile.WriteLine("<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>")
repFile.WriteLine("<center><img src='../images/b&wlogoforreport.jpg' width=300><br><br>")
repFile.WriteLine("<font size=+1><b>Inspection Report</b></font><hr noshade size=1 width='90%'></center>")
repFile.WriteLine("<table cellpadding=2 cellspacing=0 border=0 width='90%'>")
repFile.WriteLine("	<tr><td align=right><b>Date:</b></td><td colspan=3>"& Trim(rsInspec("inspecDate")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>Project Name:</b></td><td colspan=3>"& Trim(rsInspec("projectName")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right valign=top><b>Project Location:</b></td><td colspan=3 valign=top>"& Trim(rsInspec("projectAddr")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right>&nbsp;</td><td colspan=3>"& (Trim(rsInspec("projectCity")) & ", " & rsInspec("projectState") & " " & Trim(rsInspec("projectZip"))) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>County:</b></td><td colspan=3>"& Trim(rsInspec("projectCounty")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>On-Site Contact:</b></td><td colspan=3>"& Trim(rsInspec("onsiteContact")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>Office Number:</b></td><td colspan=3>"& Trim(rsInspec("officePhone")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>Emergency Number:</b></td><td colspan=3>"& Trim(rsInspec("emergencyPhone")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>Company:</b></td><td>"& Trim(rsInspec("compName")) &"</td><td align=right><b>Contact:</b></td><td>"& Trim(rsInspec("compContact")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right valign=top><b>Address:</b></td><td>"& Trim(rsInspec("compAddr")))
If Trim(rsInspec("compAddr2")) <> "" Then 
	repFile.WriteLine("<br>" & Trim(rsInspec("compAddr2"))) 
End If 
repFile.WriteLine("		</td>")
repFile.WriteLine("	<td align=right><b>Phone:</b></td><td>"& Trim(rsInspec("contactPhone")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>&nbsp;</b></td><td>"& (Trim(rsInspec("compCity")) & ", " & rsInspec("compState") & " " & Trim(rsInspec("compZip"))) &"</td><td align=right><b>Fax:</b></td><td>"& Trim(rsInspec("contactFax")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>Main Telephone Number:</b></td><td>"& Trim(rsInspec("compPhone")) &"</td><td align=right><b>E-Mail:</b></td><td>"& Trim(rsInspec("contactEmail")) &"</td></tr>")
repFile.WriteLine("	<tr><td align=right><b>Type of Report:</b></td><td>"& reportType &"</td><td align=right><b>Inches of Rain:</b></td><td>")
If reportType <> "biWeekly" Then repFile.WriteLine(inches &"</td></tr>") Else repFile.WriteLine("N/A</td></tr>") End If
repFile.WriteLine("	<tr>")
IF rsInspec("bmpsInPlace")>-1 THEN 
	repFile.WriteLine("		<td align=right><b>Are BMPs in place?</b></td><td>"& bmpsInPlace &"</td>")
END IF
IF rsInspec("sediment")>-1 THEN 
	repFile.WriteLine("		<td align=right><b>Sediment Loss or Pollution?</b></td><td>"& sediment &"</td>")
END IF 
repFile.WriteLine("	</tr>")
repFile.WriteLine("</table>")
signature = Trim(rsInspec("signature"))
rsInspec.Close
Set rsInspec = Nothing
coordSQLSELECT = "SELECT correctiveMods, coordinates, existingBMP FROM Coordinates WHERE inspecID = " & inspecID &" ORDER BY orderby"
Set rsCoord = connSWPPP.Execute(coordSQLSELECT)
repFile.WriteLine("<p><center><i>Utilizing the Site Map, SWPPP INSPECTIONS, INC. makes the following observations:</i></center><p>")
repFile.WriteLine("<table border=0 cellpadding=3 width='100%' cellspacing=0>")
If rsCoord.EOF Then
	repFile.WriteLine("<tr><td colspan='2' align='center'><i>There is no coordinate data entered at this time.</i></td></tr>")
Else
	Do While Not rsCoord.EOF
		correctiveMods = Trim(rsCoord("correctiveMods"))
		coordinates = Trim(rsCoord("coordinates"))
		existingBMP = Trim(rsCoord("existingBMP"))
		repFile.WriteLine("	<tr valign=top><td width='20%' align=right><b>Location:</b></td><td width='80%' align=left>"& coordinates &"<br></td></tr>")
		IF TRIM(rsCoord("existingBMP"))<>"-1" THEN 
			repFile.WriteLine("	<tr valign=top><td width='20%' align=right><b>Existing BMP:</b></td><td width='80%' align=left>"& existingBMP &"<br></td></tr>")
		END IF 
		repFile.WriteLine("	<tr valign=top><td width='20%' align=right><b>Corrective Modifications:</b></td><td width='80%' align=left>"& correctiveMods &"</td></tr>")
		repFile.WriteLine("	<tr><td colspan=2><hr noshade size=1 align=center width='90%'></td></tr>")
		rsCoord.MoveNext
	Loop
End If ' END No Results Found
rsCoord.Close
Set rsCoord = Nothing 
IF narrative <> "" THEN 
	repFile.WriteLine("	<TR><TD colspan=4><p>"& REPLACE(TRIM(narrative),"#@#","'") &"</TD></TR>")
END IF
repFile.WriteLine("</table><p><small>"& REPLACE(TRIM(qualifications),"#@#","'") &"</p>")
repFile.WriteLine("<p><small>I certify under penalty of law that this document and all attachments were prepared under my direction or supervision in accordance with a system designed to assure that qualified personnel properly gathered and evaluated the information submitted. Based on my inquiry of the person or persons who manage the system, or those persons directly responsible for gathering the information, the information is, to the best of my knowledge and belief, true, accurate, and complete. I am aware that there are significant penalties for submitting false information, including the possibility of fine and imprisonment for knowing violations.</small><p>")
repFile.WriteLine("<table border=0 cellpadding=2 width='100%' cellspacing=0>")
repFile.WriteLine("	<tr><td width='3%' align=left><b>Name:</b></td><td width='3%' align=left><b>Print:</b></td><td width='4%' align=left><b>Inspector:</b></td></tr>")
repFile.WriteLine("	<tr><td width='3%'><img src='../Images/signatures/"& signature &"'></td><td width='3%' align=left valign=top>"& printName &"</td><td width='4%' align=left valign=top>SWPPP INSPECTIONS, INC.</td></tr>")
repFile.WriteLine("</table><br><br></body></html>")
repFile.Close
END function %>