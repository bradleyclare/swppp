<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and Not Session("validAdmin") then Response.Redirect("../default.asp") End If
%><!-- #INCLUDE FILE="../connSWPPP.asp" --><%

Server.ScriptTimeout=1500
'Response.Write(Request.Form.Count & "<br>")
IF Request.Form.Count > 0 THEN %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Sending Email Reports</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>

<BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	marginwidth="5" marginheight="5">
<!-- #INCLUDE FILE="../adminHeader2.inc" -->
<%
	FOR EACH Item IN Request.Form
		'--	Item is ProjectsUsers.projectID &":"& Inspections.inspecID -------------------
		'--	Request(Item) is Inspections.inspecID ----------------------------------------
		'-- need to create the Email content ---------------------------------------------
		strBody=""
inspecID = Request(Item)

inspecSQLSELECT = "SELECT inspecDate, Inspections.projectID, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " &_
"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " &_
"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax, " &_
"contactEmail, reportType, inches, bmpsInPlace, sediment, " &_
"narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems, horton, hortonSignV, hortonSignLD, vscr, ldscr" &_
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
strBody=strBody &"<font size='+1'><b>Inspection Report</b></font><hr noshade size='1' ></center>"
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
strBody=strBody &"<tr>"
If rsInspec("hortonSignV") And rsInspec("vscr") <> 0 Then
	rightsSELECT = "SELECT userID, firstName, lastName, phone FROM Users WHERE userID=" & rsInspec("vscr")
	Set connRights = connSWPPP.execute(rightsSELECT)
   If Not connRights.EOF Then
	   strBody=strBody &"<td align='right'><b>VSCR:</b></td><td>" & Trim(connRights("firstName")) & " " & Trim(connRights("lastName")) & ": " & Trim(connRights("phone")) & "</td>"
	End If
End If
If rsInspec("hortonSignLD") And rsInspec("ldscr") <> 0 Then
	rightsSELECT = "SELECT userID, firstName, lastName, phone FROM Users WHERE userID=" & rsInspec("ldscr")
	Set connRights = connSWPPP.execute(rightsSELECT)
   If Not connRights.EOF Then
	   strBody=strBody &"<td align='right'><b>LDSCR:</b></td><td>" & Trim(connRights("firstName")) & " " & Trim(connRights("lastName")) & ": " & Trim(connRights("phone")) & "</td>"
	End If
End If
strBody=strBody &"</tr>"
strBody=strBody &"</table>"
signature = Trim(rsInspec("signature"))
coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)
inspecDate = Trim(rsInspec("inspecDate"))
If rsInspec("projectState") = "OK" Then 
      MsgDateStart = "11/07/2017"
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
'print dr horton questions if desired
If rsInspec("horton") Then
	'get questions
	QuestionDateStart = #12/10/2020#
	QuestionDateStart2 = #2/5/2021#
   If DateDiff("d", QuestionDateStart, inspecDate) < 1 Then
		SQLQ = "SELECT * FROM HortonQuestions WHERE orderby < 27 ORDER BY orderby"
	ElseIf DateDiff("d", QuestionDateStart2, inspecDate) < 1 Then
		SQLQ = "SELECT * FROM HortonQuestions WHERE orderby > 30 AND orderby < 57 ORDER BY orderby"
	Else
		SQLQ = "SELECT * FROM HortonQuestions WHERE orderby > 60 AND orderby < 87 ORDER BY orderby"
	End If
	Set RSQ = connSWPPP.Execute(SQLQ)
	strBody=strBody &"<hr noshade size='1' align='center' >"
	If RSQ.EOF Then
		strBody=strBody &"<p>No Questions Found</p>"
	Else
		'get answer data if available
		SQLA = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
		Set RSA = connSWPPP.execute(SQLA)
    	If RSA.EOF Then
			strBody=strBody &"<p>No Answers Found</p>"
		Else
			strBody=strBody &"<table border='0' cellpadding='3' width='100%' cellspacing='0'>"
			cnt = 0
			altColors="#ffffff"
			Do While Not RSQ.EOF
				cnt = cnt + 1
				size = "90%"
				weight = "bold"
				answer = Trim(RSA("Q"&cnt))
				If answer = Trim(RSQ("default_answer")) or answer = "na" Then
					size = "70%"
					weight = "normal"
				End If
				If answer = "na" Then
					answer = "n/a"
				End If
				strBody=strBody &"<tr bgcolor="& altColors &"><td style='font-size:"& size &"; font-weight:"& weight &"'>"& cnt & " : " & Trim(RSQ("question")) &"</td>" & _ 
					"<td style='font-size:"& size &"; font-weight:"& weight &"'>"& answer &"</td></tr>"

				If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If

				If cnt = 12 then
					pondSQL="SELECT * FROM HortonLocations WHERE inspecID="& inspecID &" AND isOutfall=0"
					'response.Write(pondSQL)
					Set RSpond=connSWPPP.execute(pondSQL)
					
					Do While Not RSpond.EOF
						locationName = Trim(RSpond("locationName")) 
						size = "90%"
						weight = "bold"
						answer = Trim(RSpond("answer"))
						If answer = "yes" Then
							size = "70%"
							weight = "normal"
						End If
						strBody=strBody &"<tr bgcolor="& altColors &"><td style='font-size:"& size &"; font-weight:"& weight &"'> - "& locationName &"</td>" & _ 
							"<td style='font-size:"& size &"; font-weight:"& weight &"'>"& answer &"</td></tr>"
						If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
						RSpond.MoveNext
            	Loop
				end if

				if cnt = 13 then
					outfallSQL="SELECT * FROM HortonLocations WHERE inspecID="& inspecID &" AND isOutfall=1"
					'response.Write(outfallSQL)
					Set RSoutfall=connSWPPP.execute(outfallSQL)

					Do While Not RSoutfall.EOF
						locationName = Trim(RSoutfall("locationName")) 
						size = "90%"
						weight = "bold"
						answer = Trim(RSoutfall("answer"))
						If answer = "yes" Then
							size = "70%"
							weight = "normal"
						End If
						strBody=strBody &"<tr bgcolor="& altColors &"><td style='font-size:"& size &"; font-weight:"& weight &"'> - "& locationName &"</td>" & _ 
							"<td style='font-size:"& size &"; font-weight:"& weight &"'>"& answer &"</td></tr>"
						If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
						RSoutfall.MoveNext
            	Loop
				end if

				RSQ.MoveNext
			Loop 'RSO
			strBody=strBody &"</table>" 
		End If
    End If
	RSQ.Close
    SET RSQ=nothing
End If
strBody=strBody &"<table border='0' cellpadding='3' width='100%' cellspacing='0'>"
strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' ></td></tr>"
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
			pond = rsCoord("pond")
			sedloss = rsCoord("sedloss")
			sedlossw = rsCoord("sedlossw")
			ce = rsCoord("ce")
			street = rsCoord("street")
			sfeb = rsCoord("sfeb")
			rockdam = rsCoord("rockdam")
			ip = rsCoord("ip")
			wo = rsCoord("wo")
			veg = rsCoord("veg")
			stock = rsCoord("stock")
			toilet = rsCoord("toilet")
			trash = rsCoord("trash")
			dewater = rsCoord("dewater")
			dust = rsCoord("dust")
        	riprap = rsCoord("riprap")
        	outfall = rsCoord("outfall")
			intop = rsCoord("intop")
			swalk = rsCoord("swalk")
			mormix = rsCoord("mormix")
			ada = rsCoord("ada")
		   dway = rsCoord("dway")
		   flume = rsCoord("flume")
			OSC = rsCoord("osc")
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
				If OSC = True Then
                correctiveMods = "(OSC) " & correctiveMods
            End If
			If pond = True Then
                correctiveMods = "(pond) " & correctiveMods
            End If
			If sedloss = True Then
                correctiveMods = "(sediment loss) " & correctiveMods
            End If
			If sedlossw = True Then
                correctiveMods = "(sediment loss to waters) " & correctiveMods
            End If
			If ce = True Then
                correctiveMods = "(construction entrance) " & correctiveMods
            End If
			If street = True Then
                correctiveMods = "(street cleaning) " & correctiveMods
            End If
			If sfeb = True Then
                correctiveMods = "(perimeter controls) " & correctiveMods
            End If
			If rockdam = True Then
	        	correctiveMods = "(rock dam) " & correctiveMods
            End If
			If ip = True Then
                correctiveMods = "(inlet protection) " & correctiveMods
            End If
			If wo = True Then
                correctiveMods = "(washout) " & correctiveMods
            End If
			If veg = True Then
                correctiveMods = "(vegetation) " & correctiveMods
            End If
			If stock = True Then
                correctiveMods = "(stockpile) " & correctiveMods
            End If
			If toilet = True Then
                correctiveMods = "(toilet) " & correctiveMods
            End If
			If trash = True Then
                correctiveMods = "(trash/waste/material) " & correctiveMods
            End If
			If dewater = True Then
				correctiveMods = "(dewatering) " & correctiveMods
			End If
			If dust = True Then
				correctiveMods = "(dust control) " & correctiveMods
			End If
			If riprap = True Then
	        	correctiveMods = "(riprap) " & correctiveMods
	      End If
	      If outfall = True Then
	        	correctiveMods = "(outfall) " & correctiveMods
	      End If
			If intop = True Then
        		correctiveMods = "(inlet top) " & correctiveMods
         End If
         If swalk = True Then
        		correctiveMods = "(sidewalk) " & correctiveMods
         End If
         If mormix = True Then
        		correctiveMods = "(mortar mix) " & correctiveMods
         End If
			If ada = True Then
				correctiveMods = "(ADA ramp) " & correctiveMods
			End If
			If dway = True Then
				correctiveMods = "(driveway) " & correctiveMods
			End If
			If flume = True Then
				correctiveMods = "(flume) " & correctiveMods
			End If
            If NLN = True Then
                'do nothing
            ElseIf infoOnly = True and (useAddress=False and coordinates="") or (useAddress=True and locationName="" and address="") Then
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
				 item_title = "action needed"
				 If infoOnly = True Then
					 item_title = "note"
				 End If
			    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>"& item_title &":</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  correctiveMods &"</td></tr>"
			    IF applyScoring and repeat THEN
				    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class = '"& scoring_class &"'>"&  age &"<br></td></tr>"
			    END IF
            End If
			strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' ></td></tr>" & vbCrLf	
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
END IF
If rsInspec("horton") Then
	projectID = rsInspec("projectID")
	projectName = Trim(rsInspec("projectName"))
	projectPhase = Trim(rsInspec("projectPhase"))
	strBody=strBody &"<br><div align='center'><a href='http://swppp.com/views/inspections.asp?projID=" & projectID & "&projName=" & projectName & "&projPhase=" & projectPhase & "'>sign off on report</a></div>"
End If
strBody=strBody &"<br><div align='center'><a href='http://www.swppp.com'>link to: www.swppp.com</a></div></Body>"

rsCoord.Close
Set rsCoord = Nothing

rsInspec.Close
Set rsInspec = Nothing

RS3.Close
SET RS3=nothing

'--	now we can create the list of recipients for the email ----------------------------------------
projID=SPLIT(Item,":")
projectID=projID(0)
		'-- Response.Write(Item &":"& Request(Item) &"<br>")
		SQL1="SELECT DISTINCT (LTRIM(RTRIM(u.firstName)) +' '+ LTRIM(RTRIM(u.lastName))) as fullName,"&_
			" u.email, u.noImages, i.projectName, i.projectPhase, i.inspecDate, pu.rights" &_
			" FROM ProjectsUsers pu JOIN Users u on pu.userID=u.userID" &_
			" JOIN Inspections i ON pu.projectID=i.projectID" &_
			" WHERE i.inspecID="& Request(Item) &" AND pu.projectID="& projectID
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		RS1.Open SQL1, connSWPPP

'--------------------- process mailing -------------------------------------------
		contentSubject= "Inspection Report for "& TRIM(RS1("projectName")) &" "& TRIM(RS1("projectPhase")) &" on "& TRIM(RS1("inspecDate"))
		Set Mailer = Server.CreateObject("Persits.MailSender")
		Mailer.FromName    = "Don Wims"
		Mailer.From        = "dwims@swppp.com"
		Mailer.Host        = "127.0.0.1"
		Mailer.Subject     = contentSubject
		Mailer.Body        = strBody & strImages & "<Body>"
		Mailer.isHTML      = True


'--------this line of code is for testing the smtp server---------------------
'		Mailer.AddBCC "SWPPP Server testing", "jzuther@gmail.com"
'--------this line of code is for testing the smtp server---------------------


'-- build the recipients list ------------------------------------------------
		fullname = TRIM(RS1("projectName")) &" "& TRIM(RS1("projectPhase"))
		DO WHILE NOT RS1.EOF
		    curRights = Trim(RS1("rights"))
            if curRights = "email" then
			    Mailer.AddAddress Trim(RS1("email")), Trim(RS1("fullName"))
			End if
            if curRights = "ecc" then
			    Mailer.AddCC Trim(RS1("email")), Trim(RS1("fullName"))
			End if
            if curRights = "bcc" then
			    Mailer.AddBCC Trim(RS1("email")), Trim(RS1("fullName"))
			End if
			RS1.MoveNext
		LOOP
		On Error Resume Next
		Mailer.Send
		If Err <> 0 Then %>
			<FONT color="red"><%=fullname%>: Mail send failure.- </FONT><%= Err.Description %><br>
<%		else %>
			<FONT color="red"><%=fullname%>: Emails Sent</FONT><br>
<%		end if
'--	now it is time to set the released bit on the inspection -------------------------------
		SQL2="UPDATE Inspections SET released=1 WHERE inspecID="& inspecID
		connSWPPP.execute(SQL2)
'--		Response.Write(strBody)
	NEXT
'	Response.End
ELSE
If Session("userID") = 1370 Then
	SQL0 = "SELECT DISTINCT i.projectName, i.projectPhase, i.inspecDate, i.inspecID, pu.projectID, i.ReportType, i.released, i.horton" &_
		" FROM ProjectsUsers pu JOIN Inspections i ON pu.projectID=i.projectID" &_
		" WHERE pu.rights='inspector' AND i.released=0 ORDER BY i.projectName, i.projectPhase, i.inspecDate DESC"
Else
	SQL0 = "SELECT DISTINCT i.projectName, i.projectPhase, i.inspecDate, i.inspecID, pu.projectID, i.ReportType, i.released, i.horton" &_
		" FROM ProjectsUsers pu JOIN Inspections i ON pu.projectID=i.projectID" &_
		" WHERE pu.rights='inspector' AND i.released=0 AND i.userID=pu.userID AND pu.userID="& Session("userID")
	If Session("userID") = 42 Then
		SQL0 = SQL0 & " and datediff(m, i.inspecdate, getdate()) <4"
	End	If
	SQL0 = SQL0 & " ORDER BY i.projectName, i.projectPhase, i.inspecDate DESC"
End If
Set RS0 = Server.CreateObject("ADODB.Recordset")
'Response.Write(SQL0)
RS0.Open SQL0, connSWPPP
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Release Reports</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>

<BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0
	marginwidth="5" marginheight="5">

<% If not Request("print") then %> <!-- #INCLUDE FILE="../adminHeader2.inc" --> <% end if %>
<h1>Send Reports via Email - <%=Session("userID")%></h1>
<FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
<div align="center">
<table border="0" cellpadding=1 cellspacing=1>
	<tr><th>Project Name|Phase</th><th>Report Date</th><th>Report Type</th><th>questions defined</th><th>send email</th></tr>
	<% DO WHILE NOT RS0.EOF 
			inspecID  = RS0("inspecID")
			projectID = RS0("projectID")
			horton    = RS0("horton")
			'Response.Write("inspecID:" & inspecID & ", horton:" & horton & "</br>")
			SQLA = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
			Set RSA = connSWPPP.execute(SQLA)
			if NOT horton Then
				question = ""
			elseif horton AND RSA.EOF Then
				question = "no"
				fc = "red"
			else
				question = "yes"
				fc = "black"
			end if %>
			<tr><td align="left"><%= RS0("projectName")%>&nbsp;<%= RS0("projectPhase") %></td>
			<td align="left"><%= RS0("inspecDate") %></td>
			<td align="left"><%= RS0("ReportType") %></td>
			<td align="center" style="color:<%=fc%>"><%=question%></td>
			<td align="center"><INPUT type="checkbox" name="<%=projectID%>:<%=inspecID%>" value="<%=inspecID%>"></td></tr>
			<% RS0.MoveNext
	LOOP
RS0.Close
SET RS0=nothing %>
</table></div>

<div align="center"><br><br>To Send These Reports via Email to all Users assigned<br />
	to Receive them and release this report <br />
    <input type="submit" value="Send Emails"></div>
</FORM>
</BODY>
</HTML><%
END IF
connSWPPP.close
SET connSWPPP=nothing %>