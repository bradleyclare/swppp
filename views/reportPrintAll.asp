<%@ Language="VBScript" %>
<% Response.buffer = false
If _
	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") _
	
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
	
End If

%><!-- #include file="../admin/connSWPPP.asp" --><%
projectID = Request("projID")
projectName = Trim(Request("projName"))
projectPhase = Trim(Request("projPhase"))
If Session("validAdmin") Then
	SQL1 = "SELECT DISTINCT inspecID, inspecDate, p.projectID" &_
		" FROM Projects as p, Inspections as i WHERE i.projectID=" & projectID &_
		" AND i.projectID=p.projectID ORDER BY inspecDate ASC"	
Else
	SQL1 = "SELECT DISTINCT inspecID, inspecDate, p.projectID" &_
		" FROM Projects as p, Inspections as i WHERE i.projectID=" & projectID &_
		" AND i.projectID=p.projectID ORDER BY inspecDate ASC"
	'SQL1 = "SELECT DISTINCT inspecID, inspecDate, p.projectName, p.projectPhase" & _
	'	" FROM Projects as p, ProjectsUsers as pu, Inspections as i" & _
	'	" WHERE pu.userID = " & Session("userID") &_
	'	" AND i.projectID=p.projectID ORDER BY inspecDate ASC"	
End If
'Response.Write(SQL1)
SET RS1 = connSWPPP.Execute(SQL1) %>
<html><head>
<title>Print All Reports <%= Date %></title>
<link rel="stylesheet" type="text/css" href="../global.css">
<style>
	.red{color: #F52006;}
	.black{color: black;}
    .bold{font-weight: bold;}
	.ldred{font-weight: bold; color: red;}
</style>
</head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15" onLoad="window.print();"><%
cnt=0
DO WHILE NOT RS1.EOF

inspecID = RS1("inspecID")
SQL2 = "SELECT inspecDate, Inspections.projectID, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " & _
	"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " & _
	"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, " & _
	"contactFax, contactEmail, reportType, inches, bmpsInPlace, " & _
	"sediment, narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems, horton, hortonSignV, hortonSignLD, vscr, ldscr, forestar" & _
	" FROM Inspections, Projects, Users" & _
	" WHERE inspecID = " & inspecID & _
	" AND Inspections.projectID = Projects.projectID" & _
	" AND Inspections.userID = Users.userID"
'Response.Write(inspecSQLSELECT)
Set RS2 = connSWPPP.Execute(SQL2)
'Response.Write("signature = " & Trim(RS2("signature")) & "<br>")
bmpsInPlace = "No"
If RS2("bmpsInPlace") = "1" Then bmpsInPlace = "Yes" End If
sediment = "No"
If RS2("sediment") ="1" Then sediment = "Yes" End If
reportType = Trim(RS2("reportType"))
inches = RS2("inches")
printName = Trim(RS2("firstName")) & " " & Trim(RS2("lastName"))
narrative= TRIM(RS2("narrative"))
IF IsNull(narrative) THEN narrative="" END IF
qualifications= TRIM(RS2("qualifications"))
IF IsNull(qualifications) THEN qualifications="" END IF %>
<DIV<% IF cnt>0 THEN %> style="page-break-before:always;"<% END IF %>>
<center><img src="../images/color_logo_report.jpg" width="300"><br><br>
<font size="+1"><b>Inspection Report</b></font>
<hr noshade size="1" width="90%"></center>
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><td align="right"><b>Date:</b></td><td colspan="3"><% = Trim(RS2("inspecDate")) %></td></tr>
	<tr><td align="right"><b>Project Name:</b></td><td colspan="3"><% = Trim(RS2("projectName")) %>&nbsp;<%= Trim(RS2("projectPhase"))%></td></tr>
	<tr><td align="right" valign="top"><b>Project Location:</b></td><td colspan="3" valign="top"><% = Trim(RS2("projectAddr")) %></td></tr>
	<tr><td align="right">&nbsp;</td><td colspan="3"><% = (Trim(RS2("projectCity")) & ", " & RS2("projectState") & " " & Trim(RS2("projectZip"))) %></td></tr>
	<tr><td align="right"><b>County:</b></td><td colspan="3"><% = Trim(RS2("projectCounty")) %></td></tr>
	<tr><td align="right"><b>On-Site Contact:</b></td><td colspan="3"><% = Trim(RS2("onsiteContact")) %></td></tr><% 
	If Len(Trim(RS2("officePhone"))) > 0 then %>
	<tr><td align="right"><b>On-Site Contact:</b></td><td colspan="3"><% = Trim(RS2("officePhone")) %></td></tr><%
	End If 
	If Len(Trim(RS2("emergencyPhone"))) > 0 then %>
	<tr><td align="right"><b>
		<% If RS2("forestar") = True Then %>
			TPDES Permit #:
		<% Else %>
			On-Site Contact
		<% End If %>
		</b></td><td colspan="3"><% = Trim(RS2("emergencyPhone")) %></td></tr><%
	End If %>
	<tr><td align="right"><b>Company:</b></td><td><% = Trim(RS2("compName")) %></td><td align="right"><b>Contact:</b></td><td><% = Trim(RS2("compContact")) %></td></tr>
	<tr><td align="right" valign="top"><b>Address:</b></td><td><% = Trim(RS2("compAddr")) %> <% If Trim(RS2("compAddr2")) <> "" Then Response.Write("<br>" & Trim(RS2("compAddr2"))) End If %></td><td align="right"><b>Phone:</b></td><td><% = Trim(RS2("contactPhone")) %></td></tr>
	<tr><td align="right"><b>&nbsp;</b></td><td><% = (Trim(RS2("compCity")) & ", " & RS2("compState") & " " & Trim(RS2("compZip"))) %></td><td align="right"><b>Fax:</b></td><td><% = Trim(RS2("contactFax")) %></td></tr>
	<tr><td align="right"><b>Main Telephone Number:</b></td><td><% = Trim(RS2("compPhone")) %></td><td align="right"><b>E-Mail:</b></td><td><% = Trim(RS2("contactEmail")) %></td></tr>
	<tr><td align="right"><b>Type of Report:</b></td><td><% = reportType %></td>
<% 	If inches="-1" THEN %>
	<td></td>
<%	Else %>
		<td align="right"><b>Inches of Rain:</b></td><td><% If reportType <> "biWeekly" Then Response.Write(inches) Else Response.Write("N/A") %></td>
<% 	End If %>
</tr>
	<tr><%  IF RS2("bmpsInPlace")>-1 THEN %><td align="right"><b>Are BMPs in place?</b></td><td><% = bmpsInPlace %></td>
<%  END IF
	IF RS2("sediment")>-1 THEN %>
		<td align="right"><b>Sediment Loss or Pollution?</b></td>
		<td><% = sediment %></td>
<% 	END IF %>
</tr>
<tr>
<% If RS2("hortonSignV") And RS2("vscr") <> 0 Then
	rightsSELECT = "SELECT userID, firstName, lastName, phone FROM Users WHERE userID=" & RS2("vscr")
	Set connRights = connSWPPP.execute(rightsSELECT)
   If Not connRights.EOF Then %>
	   <td align="right"><b>VSCR:</b></td><td> <% =Trim(connRights("firstName")) %> <% =Trim(connRights("lastName")) %>: <% =Trim(connRights("phone")) %> </td>
	<% End If
End If %>
<% If RS2("hortonSignLD") And RS2("ldscr") <> 0 Then
	rightsSELECT = "SELECT userID, firstName, lastName, phone FROM Users WHERE userID=" & RS2("ldscr")
	Set connRights = connSWPPP.execute(rightsSELECT)
   If Not connRights.EOF Then %>
	   <td align="right"><b>LDSCR:</b></td><td> <% =Trim(connRights("firstName")) %> <% =Trim(connRights("lastName")) %>: <% =Trim(connRights("phone")) %> </td>
	<% End If
End If %>
<% If RS2("forestar") And RS2("ldscr") <> 0 Then
	SQLA="SELECT * FROM HortonApprovals WHERE LD=1 and inspecID="& inspecID
	SET RSA=connSWPPP.execute(SQLA)
	forestar_SSCR_approved = false
	If not RSA.EOF Then 
		forestar_SSCR_approved = True
	End If
	RSA.Close 
	Set RSA = Nothing
	rightsSELECT = "SELECT userID, firstName, lastName, phone, signature FROM Users WHERE userID=" & RS2("ldscr")
	Set connRights = connSWPPP.execute(rightsSELECT)
    forestar_SSCR_name = ""
	forestar_SSCR_signature = ""
	If Not connRights.EOF Then 
		forestar_SSCR_name = Trim(connRights("firstName")) & " " &  Trim(connRights("lastName")) 
		forestar_SSCR_signature = Trim(connRights("signature")) %>
	   <td align="right"><b>SSCR:</b></td><td> <% =forestar_SSCR_name %>: <% =Trim(connRights("phone")) %> </td>
	<% End If
End If %>
</tr>
</table>
<% coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)%>
<table border="0" cellpadding="3" width="100%" cellspacing="0">
<tr><td align="center"><div style="font-size: 8px"><center>
<% inspecDate = Trim(RS2("inspecDate"))
If RS2("projectState") = "OK" Then
      inspecDate = Trim(RS2("inspecDate"))
      MsgDateStart = "11/07/2017"
      If DateDiff("d", inspecDate, MsgDateStart) < 1 Then %>
         <i>A qualified inspector familiar with the OPDES Permit OKR10 and the SWPPP should inspect all areas of the site that have been cleared,
         graded, or excavated and that have not yet completed stabilization; all stormwater controls (including pollution prevention measures) installed at the site; material,
         waste, borrow, or equipment storage and maintenance areas; areas where stormwater typically flows within the site, including drainage ways designed to divert, convey,
         and/or treat stormwater; all points of discharge from the site, including exit points that sediment that has been tracked out from the site; and all locations where
         stabilization measures have been implemented.<br/><br/>
         A qualified inspector should check whether all erosion and sediment controls and pollution prevention controls are properly installed, appear to be operational, and are
         working as intended to minimize pollutants discharges. Determine if any controls need to be replaced, repaired, or maintained. Check for the presence of conditions that
         could lead to spills, leaks, or other accumulations of pollutants on the site. Identify any locations where new or modified stormwater controls are necessary to minimize
         track-out, minimize dust, minimize the disturbance of steep slopes, protect storm drain inlets, and meet stabilization requirements. At discharge points and the banks of any
         surface waters, check for signs of visible erosion and sedimentation that have occurred.<br/></br/>
         Sediment must be removed before it has accumulated to one-half of the above-ground height of any perimeter control. Dewatering must have appropriate controls unless there
         is uncontaminated clear dewatering water. Cover must be provided for building materials. Chemicals must be stored in water-tight containers and provide either cover or
         secondary containment. Hazardous or toxic waste must be separated from construction or domestic waste and stored in sealed containers labeled with RCRA requirements. For
         construction and domestic waste, a dumpster or trash receptacle with a lid must be closed during rain or chance of rain, and covered at end of each work shift and when workers
         not present. Tarp or plastic must be provided if no lid is used.<br/><br/>
         If a discharge is occurring during your inspection, you are required to observe and document the visual quality of the discharge, and take note of the characteristics of the
         stormwater discharge, including color, odor, floating, settled, or suspended solids, foam, oil sheen, and other obvious indicators of stormwater pollutants; and document whether
         your stormwater controls are operating effectively, and describe any such controls that are clearly not operating as intended or are in need of maintenance.</i>
   <% Else %>
         <i>Inspectors familiar with the OPDES Permit OKR10 and the SWPPP should inspect disturbed areas of the site that have not
		 been finally stabilized, areas used for storage of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations,
		 locations where vehicles enter and exit the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas,
		 and fueling areas.</i>
   <% End If %>
<% Else 
	MsgDateStart = "3/04/2023"
	If DateDiff("d", inspecDate, MsgDateStart) < 1 Then %>
		<i>Inspectors knowledgeable of TXR150000 and the SWPPP must inspect disturbed areas, stabilization measures, construction
		  support activities, material storage areas, equipment staging areas, vehicle repair/fueling areas, locations where vehicles enter/exit the site, stockpiles of dirt, stormwater 
		  and structural controls, the discharge of pollutants, areas where stormwater flows, discharge points, and off-site areas. Report the locations where erosion and discharges of 
		  sediment or other pollutants from the site have occurred. Report locations of BMPs that need to be maintained, failed, proved inadequate, and where additional BMPs are needed.</i>
	<% Else %>
   		<i>Inspectors familiar with the TPDES Permit TXR150000 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized,
   		  areas used for storage of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles
   		  enter and exit the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i>
	<% End If %>
<% End If %>
</center></div></td></tr>
</table>
<% 'print dr horton or forestar questions if desired
If RS2("horton") or RS2("forestar") Then
	'get questions
	If RS2("horton") Then
		QuestionDateStart = #12/10/2020#
		QuestionDateStart2 = #2/5/2021#
		If DateDiff("d", QuestionDateStart, inspecDate) < 1 Then
			SQLQ = "SELECT * FROM HortonQuestions WHERE orderby < 27 ORDER BY orderby"
		ElseIf DateDiff("d", QuestionDateStart2, inspecDate) < 1 Then
			SQLQ = "SELECT * FROM HortonQuestions WHERE orderby > 30 AND orderby < 57 ORDER BY orderby"
		Else
			SQLQ = "SELECT * FROM HortonQuestions WHERE orderby > 60 AND orderby < 87 ORDER BY orderby"
		End If
	Else
		SQLQ = "SELECT * FROM HortonQuestions WHERE orderby > 90 AND orderby < 101 ORDER BY orderby"
	End If
	Set RSQ = connSWPPP.Execute(SQLQ) %>
	<hr noshade size="1" align="center" >
	<% If RSQ.EOF Then %>
		<p>No Questions Found</p>
	<% Else
		'get answer data if available
		SQLA = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
		Set RSA = connSWPPP.execute(SQLA)
    	If RSA.EOF Then %>
			<p>No Answers Found</p>
		<% Else %>
			<table border="0" cellpadding="3" width="100%" cellspacing="0">
			<% cnt = 0
			altColors="#ffffff"
			Do While Not RSQ.EOF
				cnt = cnt + 1
				size = "90%"
				weight = "bold"
				answer = Trim(RSA("Q"&cnt))
				default_val = Trim(RSQ("default_answer"))
				If default_val = "na" Then
                	default_val = "yes"
            	End If 
				If answer = default_val or answer = "na" Then
					size = "70%"
					weight = "normal"
				End If
				If answer = "na" Then
					answer = "n/a"
				End If %>
				<tr bgcolor=<%=altColors%>><td style="font-size:<%=size%>; font-weight:<% =weight %>"><%=cnt%> : <%=Trim(RSQ("question"))%></td> 
				<td style="font-size:<%=size%>; font-weight:<%=weight%>"><%=answer%></td></tr>

				<% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If

				If cnt = 12 then
					pondSQL="SELECT * FROM HortonLocations WHERE inspecID="& inspecID &" AND isOutfall=0"
					'response.Write(pondSQL)
					Set RSpond=connSWPPP.execute(pondSQL)
					
					Do While Not RSpond.EOF
						locationName = Trim(RSpond("locationName")) 
						size = "90%"
						weight = "bold"
						answer = Trim(RSpond("answer"))
						If answer = "yes" or answer = "na" Then
							size = "70%"
							weight = "normal"
						End If 
						If answer = "na" Then
							answer = "n/a"
						End If %>
						<tr bgcolor=<%=altColors%>><td style="font-size:<%=size%>; font-weight:<% =weight %>">&nbsp-&nbsp<% = locationName%></td> 
						<td style="font-size:<%=size%>; font-weight:<%=weight%>"><%=answer%></td></tr>
						<% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
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
						If answer = "yes" or answer = "na" Then
							size = "70%"
							weight = "normal"
						End If
						If answer = "na" Then
							answer = "n/a"
						End If %>
						<tr bgcolor=<%=altColors%>><td style="font-size:<%=size%>; font-weight:<% =weight %>">&nbsp-&nbsp<% = locationName%></td> 
						<td style="font-size:<%=size%>; font-weight:<%=weight%>"><%=answer%></td></tr>
						<% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
						RSoutfall.MoveNext
            	Loop
				end if

				RSQ.MoveNext
			Loop %>
			</table> 
			<hr noshade size="1" align="center" >
		<% End If
    End If
	RSQ.Close
    SET RSQ=nothing
End If %>
<table border="0" cellpadding="3" width="100%" cellspacing="0">
<% 
If RS2("compliance") Then
	Response.Write("<tr><td colspan='2' align='center'><h2>SITE IS IN COMPLIANCE WITH THE SWPPP</h2></td></tr>")
Else 
    If rsCoord.EOF Then
	    Response.Write("<tr><td colspan='2' align='center'><i>There is no " & _
		"coordinate data entered at this time.</i></td></tr>")
    Else
		If RS2("forestar") Then
			applyScoring = true
		else
			applyScoring = false
		End If
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
			dis = rsCoord("discharge")
			scoring_class = "black"
			IF applyScoring and repeat THEN
				IF assignDate = "" THEN
					age = 0
				ELSE
					age = datediff("d",assignDate,currentDate) 
				END IF
				'Response.Write("Age:" & age & " - " & currentDate & "</br>")
				IF age >= 7 THEN
					If RS2("forestar") Then
						scoring_class = "red"
					Else
						scoring_class = "bold"
					End If
				END IF
			END IF
			'Response.Write("Scoring Class:" & scoring_class & "</br>")
			'Response.Write("ID: " & coID & ", Repeat: " & repeat & ", Age: " & age & ", Coord: " & coordinates & ", LocName: " & locationName & ", address: " & address & ", Mods: " & correctiveMods & "<br/>") 
			If LD = True Then
               correctiveMods = "(LD) " & correctiveMods
			   If RS2("forestar") and repeat Then
			   	  scoring_class = "ldred"
			   Else
                  scoring_class = "bold"
			   End If
            End If
			If OSC = True Then
               correctiveMods = "(OSC) " & correctiveMods
            End If
			If 0 Then 'Session("validInspector") Then
				If wo = True Then
					correctiveMods = "(wo) " & correctiveMods
				End If
				If veg = True Then
					correctiveMods = "(veg) " & correctiveMods
				End If
				If toilet = True Then
					correctiveMods = "(toilet) " & correctiveMods
				End If
				If trash = True Then
					correctiveMods = "(trash) " & correctiveMods
				End If
				If swalk = True Then
					correctiveMods = "(swalk) " & correctiveMods
				End If
				If stock = True Then
					correctiveMods = "(stock) " & correctiveMods
				End If
				If street = True Then
					correctiveMods = "(street) " & correctiveMods
				End If
				If sedlossw = True Then
					correctiveMods = "(sedlossw) " & correctiveMods
				End If
				If sedloss = True Then
					correctiveMods = "(sedloss) " & correctiveMods
				End If
				If rockdam = True Then
					correctiveMods = "(rockdam) " & correctiveMods
				End If
				If riprap = True Then
					correctiveMods = "(riprap) " & correctiveMods
				End If
				If pond = True Then
					correctiveMods = "(pond) " & correctiveMods
				End If
				If outfall = True Then
					correctiveMods = "(outfall) " & correctiveMods
				End If
				If mormix = True Then
					correctiveMods = "(mormix) " & correctiveMods
				End If
				If ip = True Then
					correctiveMods = "(ip) " & correctiveMods
				End If
				If intop = True Then
					correctiveMods = "(intop) " & correctiveMods
				End If
				If flume = True Then
					correctiveMods = "(flume) " & correctiveMods
				End If
				If sfeb = True Then
					correctiveMods = "(eb/tf/sf) " & correctiveMods
				End If
				If dway = True Then
					correctiveMods = "(dway) " & correctiveMods
				End If
				If dust = True Then
					correctiveMods = "(dust) " & correctiveMods
				End If
				If dis = True Then
					correctiveMods = "(dis) " & correctiveMods
				End If
				If dewater = True Then
					correctiveMods = "(dewater) " & correctiveMods
				End If
				If ce = True Then
					correctiveMods = "(ce) " & correctiveMods
				End If
				If ada = True Then
					correctiveMods = "(ada) " & correctiveMods
				End If
			End If
            If infoOnly = True and (useAddress=False and coordinates="") or (useAddress=True and locationName="" and address="") Then %>
                <tr valign='top'><td width='20%' align='right'><b>note:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=correctiveMods%></td></tr>
            <% Else
		        IF useAddress THEN %>
			        <tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class = '<%=scoring_class%>'><%=locationName%><br></td></tr>
			        <tr valign='top'><td width='20%' align='right'><b>address:</b></td>	<td width='80%' align='left' class = '<%=scoring_class%>'><%=address%><br></td></tr>
		        <% ELSE %>
			        <tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class = '<%=scoring_class%>'><%=coordinates%><br></td></tr>
		        <% END IF
		        IF TRIM(rsCoord("existingBMP"))<>"-1" THEN %>
			        <tr valign='top'><td width='20%' align='right'><b>existing BMP:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=existingBMP%><br></td></tr>
		        <% END IF  
				  item_title = "action needed"
				  If infoOnly = True Then
						item_title = "note"
				  End If %>
		        <tr valign='top'><td width='20%' align='right'><b><%=item_title%>:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=correctiveMods%></td></tr>
		        <% IF applyScoring and repeat THEN %>
			        <tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=age%> days<br></td></tr>
		        <% END IF
            End If %>
		    <tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>
		    <% rsCoord.MoveNext
	    Loop
    End If ' END No Results Found
End If 'END compliance
%>

</table>

<%
signature = Trim(RS2("signature"))
%> 

<table border="0" cellpadding="3" width="100%" cellspacing="0">
<tr><td align="center"><div style="font-size: 7px">
<% If RS2("projectState") = "OK" Then %>
	You must initiate stabilization measures immediately whenever earth-disturbing 
	activities have permanently or temporarily ceased on any portion of the site and 
	will not resume for a period exceeding 14 calendar days.
<% Else %>
	Erosion control and stabilization measures must be initiated immediately 
	in portions of the site where construction activities have temporarily ceased and 
	will not resume for a period exceeding 14 calendar days. Stabilization measures 
	that provide a protective cover must be initiated immediately in portions of the 
	site where construction activities have permanently ceased.
<% END IF %>
</div></td></tr>
<tr><td align="center"><div style="font-size: 7px"><%= REPLACE(TRIM(qualifications),"#@#","'")%></div></td></tr>
</table>
<table border="0" cellpadding="3" width="100%" cellspacing="0">
<tr><td align="center"><div style="font-size: 7px">
I certify under penalty of law that this document and all attachments 
	were prepared under my direction or supervision in accordance with a system 
	designed to assure that qualified personnel properly gathered and evaluated 
	the information submitted. Based on my inquiry of the person or persons who 
	manage the system, or those persons directly responsible for gathering the 
	information, the information is, to the best of my knowledge and belief, true, 
	accurate, and complete. I am aware that there are significant penalties for 
	submitting false information, including the possibility of fine and imprisonment 
	for knowing violations.
</div></td></tr>
</table>
<br/>
<table border="0" cellpadding="3" width="100%" cellspacing="0">
	<tr><td width="3%" align="left"><b>Name:</b></td><td width="3%" align="left"><b>Print:</b></td><td width="4%" align="left"><b>Inspector:</b></td></tr>
	<tr><td width="3%"><img src="../images/signatures/<% = signature %>"></td><td width="3%" align="left" valign="top"><% = printName %></td><td width="4%" align="left" valign="top">SWPPP INSPECTIONS, INC.</td></tr>
</table>

<%	If RS2("forestar") and forestar_SSCR_approved Then  %>
	<p><div style="font-size: 7px">
	I certify under penalty of law that this document and all attachments were prepared under 
	my direction or supervision in accordance with a system designed to assure that qualified 
	personnel properly gathered and evaluated the information submitted. Based on my inquiry 
	of the person or persons who manage the system, or those persons directly responsible for 
	gathering the information, the information is, to the best of my knowledge and belief, true, 
	accurate, and complete. I am aware that there are significant penalties for submitting false 
	information, including the possibility of fine and imprisonment for knowing violations.
	</div><p> 
	<table border="0" cellpadding="2" width="100%" cellspacing="0">
		<tr> 
			<td width="3%" align="left"><b>Name:</b></td>
			<td width="3%" align="left"><b>Site Stormwater Compliance Representative (SSCR):</b></td>
		</tr>
		<tr> 
			<td width="3%"><img src="../images/signatures/<% = forestar_SSCR_signature %>"></td>
			<td width="3%" align="left" valign="top"><% = forestar_SSCR_name %></td>
		</tr>
	</table>
<% End If %>

<br><br></DIV><%
	RS1.MoveNext 
	cnt=1
LOOP
RS1.Close
SET RS1=nothing

rsCoord.Close
Set rsCoord = Nothing 

RS2.Close
Set RS2 = Nothing 
%>
</body>
</html><%
connSWPPP.Close
Set connSWPPP = Nothing %>