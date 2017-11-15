<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and Not Session("validAdmin") then Response.Redirect("../default.asp") End If %>
<!-- #INCLUDE FILE="../connSWPPP.asp" -->
<% Server.ScriptTimeout=1500
'Response.Write(Request.Form.Count & "<br>") 
%>
<!-- #INCLUDE FILE="../adminHeader2.inc" -->
<% strBody=""
inspecID = Request("inspecID")

inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " &_
"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " &_
"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax, " &_
"contactEmail, reportType, inches, bmpsInPlace, sediment, " &_
"narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems" &_
" FROM Inspections, Projects, Users" &_
" WHERE inspecID = " & inspecID &_
" AND Inspections.projectID = Projects.projectID" &_
" AND Inspections.userID = Users.userID"
'Response.Write("Inspec: "& inspecSQLSELECT &"<br>")

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
strBody=strBody &"<head><style>"
strBody=strBody &".red{color: #F52006;}"
strBody=strBody &".black{color: black;}"
strBody=strBody &".ld{font-weight: bold;}"
strBody=strBody &"</style></head>"
strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
strBody=strBody &"<center><img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='300'><br><br>"
strBody=strBody &"<font size='+1'><b>Inspection Report</b></font><hr noshade size='1' width='90%'></center>"
strBody=strBody &"<table cellpadding='2' cellspacing='0' border='0' width='90%'>"
strBody=strBody &"<tr><td align='right'><b>Date:</b></td><td colspan='3'>"&  Trim(rsInspec("inspecDate")) &"</td></tr>"
strBody=strBody &"<tr><td align='right'><b>Project Name:</b></td><td colspan='3'>"&  Trim(rsInspec("projectName")) &"&nbsp;"&  Trim(rsInspec("projectPhase")) &"</td></tr>"
strBody=strBody &"<tr><td align='right' valign='top'><b>Project Location:</b></td><td colspan='3' valign='top'>"&  Trim(rsInspec("projectAddr")) &"</td></tr>"
strBody=strBody &"<tr><td align='right'>&nbsp;</td><td colspan='3'>"&  (Trim(rsInspec("projectCity")) &", "& rsInspec("projectState") &" "& Trim(rsInspec("projectZip"))) &"</td></tr>"
strBody=strBody &"<tr><td align='right'><b>County:</b></td><td colspan='3'>"&  Trim(rsInspec("projectCounty")) &"</td></tr>"
strBody=strBody &"<tr><td align='right'><b>On-Site Contact:</b></td><td colspan='3'>"&  Trim(rsInspec("onsiteContact")) &"</td></tr>"
if Trim(rsInspec("officePhone")) <> "" Then
	strBody=strBody &"<tr><td align='right'><b>On-Site Contact:</b></td><td colspan='3'>"&  Trim(rsInspec("officePhone")) &"</td></tr>"
End If
if Trim(rsInspec("emergencyPhone")) <> "" Then
	strBody=strBody &"<tr><td align='right'><b>On-Site Contact:</b></td><td colspan='3'>"&  Trim(rsInspec("emergencyPhone")) &"</td></tr>"
End If
strBody=strBody &"<tr><td align='right'><b>Company:</b></td><td>"&  Trim(rsInspec("compName")) &"</td><td align='right'><b>Contact:</b></td><td>"&  Trim(rsInspec("compContact")) &"</td></tr>"
strBody=strBody &"<tr><td align='right' valign='top'><b>Address:</b></td><td>"&  Trim(rsInspec("compAddr"))
If Trim(rsInspec("compAddr2")) <> "" Then
	strBody=strBody &"<br>"& Trim(rsInspec("compAddr2"))
End If
strBody=strBody &"</td><td align='right'><b>Phone:</b></td><td>"&  Trim(rsInspec("contactPhone")) &"</td></tr>"
strBody=strBody &"<tr><td align='right'><b>&nbsp;</b></td><td>"&  (Trim(rsInspec("compCity")) &", "& rsInspec("compState") &" "& Trim(rsInspec("compZip"))) &"</td><td align='right'><b>Fax:</b></td><td>"&  Trim(rsInspec("contactFax")) &"</td></tr>"
strBody=strBody &"<tr><td align='right'><b>Main Telephone Number:</b></td><td>"&  Trim(rsInspec("compPhone")) &"</td><td align='right'><b>E-Mail:</b></td><td>"&  Trim(rsInspec("contactEmail")) &"</td></tr>"
strBody=strBody &"<tr><td align='right'><b>Type of Report:</b></td><td>"&  reportType &"</td>"
IF inches>-1 THEN
	strBody=strBody &"<td align='right'><b>Inches of Rain:</b></td><td>"
	If reportType <> "biWeekly" Then
		strBody=strBody & inches & "</td>"
	Else
		strBody=strBody &"N/A</td>"
	End If
ELSE
	strBody=strBody &"<td></td>"
END IF
strBody=strBody &"</tr><tr>"
IF rsInspec("bmpsInPlace")>-1 THEN
	strBody=strBody &"<td align='right'><b>Are BMPs in place?</b></td><td>"&  bmpsInPlace &"</td>"
END IF
IF rsInspec("sediment")>-1 THEN
	strBody=strBody &"<td align='right'><b>Sediment Loss or Pollution?</b></td><td>"&  sediment &"</td>"
END IF
strBody=strBody &"</tr>"
strBody=strBody &"</table>"
signature = Trim(rsInspec("signature"))
coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD, NLN" &_
	" FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)
If rsInspec("projectState") = "OK" Then
	   inspecDate = Cdate(Trim(rsInspec("inspecDate")))
      MsgDateStart = #11/07/2017#
      If DateDiff("d", inspecDate, MsgDateStart) < 1 Then
         strBody=strBody &"<div style='font-size: 8px'><p><center><i>A qualified inspector familiar with the OPDES Permit OKR10 and the SWPPP should inspect all areas of the site that have been cleared," &_
         " graded, or excavated and that have not yet completed stabilization; all stormwater controls (including pollution prevention measures) installed at the site; material," &_
         " waste, borrow, or equipment storage and maintenance areas; areas where stormwater typically flows within the site, including drainage ways designed to divert, convey," &_
         " and/or treat stormwater; all points of discharge from the site, including exit points that sediment that has been tracked out from the site; and all locations where" &_
         " stabilization measures have been implemented.<br/><br/>" &_
         "A qualified inspector should check whether all erosion and sediment controls and pollution prevention controls are properly installed, appear to be operational, and are" &_
         " working as intended to minimize pollutants discharges. Determine if any controls need to be replaced, repaired, or maintained. Check for the presence of conditions that" &_
         " could lead to spills, leaks, or other accumulations of pollutants on the site. Identify any locations where new or modified stormwater controls are necessary to minimize" &_
         " track-out, minimize dust, minimize the disturbance of steep slopes, protect storm drain inlets, and meet stabilization requirements. At discharge points and the banks of any" &_
         " surface waters, check for signs of visible erosion and sedimentation that have occurred.<br/></br/>" &_
         "Sediment must be removed before it has accumulated to one-half of the above-ground height of any perimeter control. Dewatering must have appropriate controls unless there" &_
         " is uncontaminated clear dewatering water. Cover must be provided for building materials. Chemicals must be stored in water-tight containers and provide either cover or" &_
         " secondary containment. Hazardous or toxic waste must be separated from construction or domestic waste and stored in sealed containers labeled with RCRA requirements. For" &_
         " construction and domestic waste, a dumpster or trash receptacle with a lid must be closed during rain or chance of rain, and covered at end of each work shift and when workers" &_
         " not present. Tarp or plastic must be provided if no lid is used.<br/><br/>" &_
         "If a discharge is occurring during your inspection, you are required to observe and document the visual quality of the discharge, and take note of the characteristics of the" &_
         " stormwater discharge, including color, odor, floating, settled, or suspended solids, foam, oil sheen, and other obvious indicators of stormwater pollutants; and document whether" &_
         " your stormwater controls are operating effectively, and describe any such controls that are clearly not operating as intended or are in need of maintenance.</i></center></p></div>"
   Else
         strBody=strBody &"<div style='font-size: 8px'><p><center><i>Inspectors familiar with the OPDES Permit OKR10 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized, areas used for storage" &_
         " of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles enter and exit" &_
         " the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i></center></p></div>"
   End If
Else
	strBody=strBody &"<div style='font-size: 8px'><p><center><i>Inspectors familiar with the TPDES Permit TXR150000 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized," &_
      " areas used for storage of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles" &_
      " enter and exit the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i></center></p></div>"
End If
strBody=strBody &"<table border='0' cellpadding='3' width='100%' cellspacing='0'>"
strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"
If rsInspec("compliance") Then
	strBody=strBody &"<tr><td colspan='2' align='center'><h2>SITE IS IN COMPLIANCE WITH THE SWPPP</h2></td></tr>"
Else 
	If rsCoord.EOF Then
		strBody=strBody &"<tr><td colspan='2' align='center'><i>There is no coordinate data entered at this time.</i></td></tr>"
	Else
		applyScoring = False 'rsInspec("includeItems")
		currentDate = date()
		Do While Not rsCoord.EOF
			coID = rsCoord("coID")
			correctiveMods = Trim(rsCoord("correctiveMods"))
			orderby = rsCoord("orderby")
			coordinates = Trim(rsCoord("coordinates"))
			existingBMP = Trim(rsCoord("existingBMP")) 
			assignDate = rsCoord("assignDate") 
			completeDate = rsCoord("completeDate")
			status = rsCoord("status")
			repeat = rsCoord("repeat")
			useAddress = rsCoord("useAddress")
			address = TRIM(rsCoord("address"))
			locationName = TRIM(rsCoord("locationName"))
            infoOnly = rsCoord("infoOnly")
            LD = rsCoord("LD")
            NLN = rsCoord("NLN")
			scoring_class = "black"
			'Response.Write("ID: " & coID & ", Coord: " & coordinates & ", LocName: " & locationName & ", address: " & address & ", Mods: " & correctiveMods & "<br/>") 
			IF applyScoring THEN
				IF assignDate = "" THEN
					age = 0
				ELSE
					age = datediff("d",assignDate,currentDate) 
				END IF
				IF age > 7 THEN
					scoring_class = "red"
				END IF
			END IF
            If LD = True Then
                correctiveMods = "(LD) " & correctiveMods
                scoring_class = "ld"
            End If
            If NLN = True Then
                'do nothing
            ElseIf infoOnly = True Then
			    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>note:</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  correctiveMods &"</td></tr>"
            Else
                IF useAddress THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class = '"& scoring_class &"'>"&  locationName &"<br></td></tr>"
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>address:</b></td>	<td width='80%' align='left' class = '"& scoring_class &"'>"&  address &"<br></td></tr>"
			    ELSE
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class = '"& scoring_class &"'>"&  coordinates &"<br></td></tr>"
			    END IF
			    IF TRIM(rsCoord("existingBMP"))<>"-1" THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>existing BMP:</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  existingBMP &"<br></td></tr>"
			    END IF
			    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>action needed:</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  correctiveMods &"</td></tr>"
			    IF applyScoring and repeat THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  age &"<br></td></tr>"
			    END IF
            End If
			strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>" & vbCrLf	
        rsCoord.MoveNext
		Loop
	End If ' END No Results Found
End If
If rsInspec("projectState") = "OK" Then
	strBody=strBody &"<TR><TD colspan=4><p><small>You must initiate stabilization measures immediately whenever earth-disturbing activities have permanently or temporarily ceased on any portion of the site and will not resume for a period exceeding 14 calendar days.</small></TD></TR>"
Else
	strBody=strBody &"<TR><TD colspan=4><p><small>Erosion control and stabilization measures must be initiated immediately in portions of the site where construction activities have temporarily ceased and will not resume for a period exceeding 14 calendar days. Stabilization measures that provide a protective cover must be initiated immediately in portions of the site where construction activities have permanently ceased.</small></TD></TR>"
END IF
strBody=strBody &"</table><p><small>"& REPLACE(TRIM(qualifications),"#@#","'")&"</small></p>"
strBody=strBody &"<p><small>I certify under penalty of law that this document and all attachments were prepared under my direction or supervision in accordance with a system designed to assure that qualified personnel properly gathered and evaluated the information submitted. Based on my inquiry of the person or persons who manage the system, or those persons directly responsible for gathering the information, the information is, to the best of my knowledge and belief, true, accurate, and complete. I am aware that there are significant penalties for submitting false information, including the possibility of fine and imprisonment for knowing violations.</small><p><table border='0' cellpadding='2' width='100%' cellspacing='0'>"
strBody=strBody &"<tr><td width='3%' align='left'><b>Name:</b></td><td width='3%' align='left'><b>Print:</b></td><td width='4%' align='left'><b>Inspector:</b></td></tr>"
strBody=strBody &"<tr><td width='3%'><img src='http://www.swpppinspections.com/images/signatures/"&  signature &"'></td><td width='3%' align='left' valign='top'>"&  printName &"</td><td width='4%' align='left' valign='top'>SWPPP INSPECTIONS, INC.</td></tr></table>"
strBody=strBody &"<br><br>"
SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& inspecID
SET RS3=connSWPPP.execute(SQL3)
IF NOT(RS3.EOF) THEN
	strBody=strBody &"<div align='center'><a href='http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName")) &"'>link for Site Map</a></div>"
END IF

'-- images portion -----------------------------------------------------------------------------------
imgSQLSELECT = "SELECT imageID, largeImage, smallImage, description FROM Images WHERE inspecID = " & inspecID
'--Response.Write("images?: "& imgSQLSELECT &"<br>")
Set rsImages = connSWPPP.execute(imgSQLSELECT)
strImages=""
If Not rsImages.EOF Then
	strImages="<div class=indent30><table cellspacing=0 cellpadding=4 width='90%' border=0><tr>"
	Do While Not rsImages.EOF
		iDataRows = iDataRows + 1
		If iDataRows > 3 Then
			strImages=strImages &"</tr>"& VBCrLf &"<tr>"
			iDataRows = 1
		End If
		strImages=strImages &"<td align=center><a href='"& "http://www.swpppinspections.com/images/lg/" & Trim(rsImages("largeImage")) &"' target='_blank'>"& Trim(rsImages("description")) &"<br>"
		If Right(Trim(rsImages("smallImage")),3)="pdf" then
			strImages=strImages &"<img src='http://www.swpppinspections.com/images/acrobat.gif' width=87 height=30 border=0 alt='Acrobat PDF Doc'>"
		else
			strImages=strImages &"<img src='"& "http://www.swpppinspections.com/images/sm/" & Trim(rsImages("smallImage")) &"' border=0	alt='"& Trim(rsImages("smallImage")) &"'>"
		end if
		strImages=strImages &"</a></td>"
		rsImages.MoveNext
	Loop
End If
strBody=strBody &"<br><div align='center'><a href='http://www.swppp.com'>link to: www.swppp.com</a></div></Body>"

rsCoord.Close
Set rsCoord = Nothing

rsInspec.Close
Set rsInspec = Nothing

RS3.Close
SET RS3=nothing

'--	now we can create the list of recipients for the email ----------------------------------------
projectID = Request("projID")
'-- Response.Write(Item &":"& Request(Item) &"<br>")
SQL1="SELECT DISTINCT (LTRIM(RTRIM(u.firstName)) +' '+ LTRIM(RTRIM(u.lastName))) as fullName,"&_
	" u.email, u.noImages, i.projectName, i.projectPhase, i.inspecDate, pu.rights" &_
	" FROM ProjectsUsers pu JOIN Users u on pu.userID=u.userID" &_
	" JOIN Inspections i ON pu.projectID=i.projectID" &_
	" WHERE i.inspecID="& inspecID &" AND pu.projectID="& projectID
Set RS1 = Server.CreateObject("ADODB.Recordset")
RS1.Open SQL1, connSWPPP

'--------------------- process mailing -------------------------------------------
contentSubject= "Inspection Report for "& TRIM(RS1("projectName")) &" "& TRIM(RS1("projectPhase")) &" on "& TRIM(RS1("inspecDate"))
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName    = "SWPPP.COM"
Mailer.FromAddress = "dwims@swppp.com"
Mailer.RemoteHost = "127.0.0.1"
Mailer.Subject    = contentSubject
Mailer.BodyText = strBody & strImages & "<Body>"
BodyText = strBody & strImages & "<Body>"
Mailer.ContentType = "text/html"
Mailer.AddRecipient "Brad Leishman", "bradleyclare@gmail.com" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Test Release Reports</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>
<BODY>
<h1>Report to be Sent</h1>
<h3>SUBJECT: <%=Mailer.Subject%></h3>
<h3>BODY:</h3>
<%=BodyText%>
</BODY>
</HTML>
<% connSWPPP.close
SET connSWPPP=nothing %>