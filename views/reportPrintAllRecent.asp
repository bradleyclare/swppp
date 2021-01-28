<%@ Language="VBScript" %>
<%Response.Buffer = False%>
<%
If _
	Not Session("validAdmin") And _
	Not Session("validUser") And _
	Not Session("validDirector") And _
	Not Session("validInspector") _
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
	
End If

%><!-- #include file="../admin/connSWPPP.asp" --><%
projectCounty = Request("cnty")
If Session("validAdmin") Then
'--	SQL1="SELECT MAX(inspecDate) as inspecDate, projectName, projectPhase FROM Inspections " &_
'--		" GROUP BY projectName, projectPhase ORDER BY projectName, projectPhase"
	SQL1="SELECT MAX(inspecDate) as inspecDate, projectid, projectname FROM Inspections GROUP BY projectid, projectname ORDER BY projectid"
Else
'--	SQL1="SELECT MAX(inspecDate) as inspecDate, projectName, projectPhase FROM Inspections i, ProjectsUsers pu " &_
'--		" WHERE pu.userID = " & Session("userID") &" AND pu.projectID=i.projectID GROUP BY projectName, projectPhase ORDER BY projectName, projectPhase"
	SQL1="SELECT MAX(inspecDate) as inspecDate, i.projectid, i.projectname FROM Inspections i, ProjectsUsers pu " &_
		" WHERE pu.userID = " & Session("userID") &" AND pu.projectID=i.projectID GROUP BY i.projectid, i.projectname ORDER BY i.projectname, i.projectid"
End If
SET RS1 = connSWPPP.Execute(SQL1) 
'--Response.Write(SQL1 &"<br><br>")
%>
<html><head>
<title>Print All Recent Reports <%= Date %></title>
<link rel="stylesheet" type="text/css" href="../global.css">
<style>
	.red{color: #F52006;}
	.black{color: black;}
    .ld{font-weight: bold;}
</style>
</head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15" onLoad="window.print();"><% '--window.close();
DO WHILE NOT RS1.EOF

    inspecDate = RS1("inspecDate")
'--projectName = RS1("projectName")
'--projectPhase= RS1("projectPhase")
    projectID = RS1("projectid")

SQL2 = "SELECT inspecID, inspecDate, Inspections.projectID, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " & _
	"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " & _
	"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, " & _
	"contactFax, contactEmail, reportType, inches, bmpsInPlace, " & _
	"sediment, narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems, horton, hortonSignV, hortonSignLD, vscr, ldscr" & _
	" FROM Inspections, Projects, Users" & _
	" WHERE inspecDate = '" & inspecDate &"' AND Inspections.projectid = "& projectID
'--	" WHERE inspecDate = '" & inspecDate &"' AND Inspections.projectName='"&  Replace(projectName, "'", "''") &"'" 
'--IF IsNull(projectPhase) THEN 
'--	SQL2=SQL2 & " AND Inspections.projectPhase is null" 
'--ELSE 
'--	SQL2=SQL2 &" AND Inspections.projectPhase='"& projectPhase &"'" 
'--END IF
	SQL2=SQL2 &" AND Inspections.projectID = Projects.projectID AND Inspections.userID = Users.userID"

'Response.Write(SQL2 &"<br><br>")

Set RS2 = connSWPPP.Execute(SQL2)
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
<DIV style="page-break-after:always;">
<% 	IF NOT(Session("validAdmin") OR Session("validInspector")) THEN %>
<center><img src="../images/color_logo_report.jpg" width="300"><br><br>
<%	End IF %>
<font size="+1"><b>Inspection Report</b></font><hr noshade size="1" width="90%"></center>
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
	<tr><td align="right"><b>On-Site Contact:</b></td><td colspan="3"><% = Trim(RS2("emergencyPhone")) %></td></tr><%
	End If %>
	<tr><td align="right"><b>Company:</b></td><td><% = Trim(RS2("compName")) %></td><td align="right"><b>Contact:</b></td><td><% = Trim(RS2("compContact")) %></td></tr>
	<tr><td align="right" valign="top"><b>Address:</b></td><td><% = Trim(RS2("compAddr")) %> <% If Trim(RS2("compAddr2")) <> "" Then Response.Write("<br>" & Trim(RS2("compAddr2"))) End If %></td><td align="right"><b>Phone:</b></td><td><% = Trim(RS2("contactPhone")) %></td></tr>
	<tr><td align="right"><b>&nbsp;</b></td><td><% = (Trim(RS2("compCity")) & ", " & RS2("compState") & " " & Trim(RS2("compZip"))) %></td><td align="right"><b>Fax:</b></td><td><% = Trim(RS2("contactFax")) %></td></tr>
	<tr><td align="right"><b>Main Telephone Number:</b></td><td><% = Trim(RS2("compPhone")) %></td><td align="right"><b>E-Mail:</b></td><td><% = Trim(RS2("contactEmail")) %></td></tr>
	<tr><td align="right"><b>Type of Report:</b></td><td><% = reportType %></td>
<% 	If inches="-1" THEN %>
	<td></td>
<%	Else %>
<td align="right"><b>Inches of Rain:</b></td><td><% If reportType <> "biWeekly" Then Response.Write(inches) Else Response.Write("N/A") %></td></tr>
<%	End If %>
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
</tr>
</table><%
inspecID= RS2("inspecID")
signature = Trim(RS2("signature"))
coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)%>
<p><center><div style="font-size: 8px">
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
         <i>Inspectors familiar with the OPDES Permit OKR10 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized, areas used for storage 
         of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles enter and exit 
         the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i>
   <% End If %>
<% Else %>
   <i>Inspectors familiar with the TPDES Permit TXR150000 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized,
   areas used for storage of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles
   enter and exit the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i>
<% End If %>
</div></center></p>
<% 'print dr horton questions if desired
If RS2("horton") Then
	'get questions
	QuestionDateStart = #12/10/2020#
   If DateDiff("d", QuestionDateStart, inspecDate) < 1 Then
		SQLQ = "SELECT * FROM HortonQuestions WHERE orderby < 27 ORDER BY orderby"
	Else
		SQLQ = "SELECT * FROM HortonQuestions WHERE orderby > 30 AND orderby < 57 ORDER BY orderby"
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
				If answer = Trim(RSQ("default_answer")) or answer = "na" Then
					size = "70%"
					weight = "normal"
				End If
				If answer = "na" Then
					answer = "n/a"
				End If %>
				<tr bgcolor=<%=altColors%>><td style="font-size:<%=size%>; font-weight:<% =weight %>"><%=cnt%> : <%=Trim(RSQ("question"))%></td> 
				<td style="font-size:<%=size%>; font-weight:<%=weight%>"><%=answer%></td></tr>
				<% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
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
	    applyScoring = False
	    'if RS2("includeItems")=True & Session("seeScoring")=True Then applyScoring = True End If
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
            ElseIf infoOnly = True and (useAddress=False and coordinates="") or (useAddress=True and locationName="" and address="") Then %>
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
			        <tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=age%><br></td></tr>
		        <% END IF
            End If %>
		    <tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>
		    <% rsCoord.MoveNext
	    Loop
    End If ' END No Results Found
End If ' END Compliance
rsCoord.Close
Set rsCoord = Nothing %>
</table>
<%	
If RS2("projectState") = "OK" Then %>
	<p><div style="font-size: 10px">
	You must initiate stabilization measures immediately whenever earth-disturbing 
	activities have permanently or temporarily ceased on any portion of the site and 
	will not resume for a period exceeding 14 calendar days.
	</div></p>
<% Else %>
	<p><div style="font-size: 10px">Erosion control and stabilization measures must be initiated immediately 
	in portions of the site where construction activities have temporarily ceased and 
	will not resume for a period exceeding 14 calendar days. Stabilization measures 
	that provide a protective cover must be initiated immediately in portions of the 
	site where construction activities have permanently ceased.</div></p>
<% END IF %>
<% 	IF NOT(Session("validAdmin") OR Session("validInspector")) THEN %>
<p><div style="font-size: 10px"><%= REPLACE(TRIM(qualifications),"#@#","'")%></div></p>
<p><div style="font-size: 10px">I certify under penalty of law that this document and all attachments 
	were prepared under my direction or supervision in accordance with a system 
	designed to assure that qualified personnel properly gathered and evaluated 
	the information submitted. Based on my inquiry of the person or persons who 
	manage the system, or those persons directly responsible for gathering the 
	information, the information is, to the best of my knowledge and belief, true, 
	accurate, and complete. I am aware that there are significant penalties for 
	submitting false information, including the possibility of fine and imprisonment 
	for knowing violations.</div></p> 
<table border="0" cellpadding="2" width="100%" cellspacing="0">
	<tr><td width="3%" align="left"><b>Name:</b></td><td width="3%" align="left"><b>Print:</b></td><td width="4%" align="left"><b>Inspector:</b></td></tr>
	<tr><td width="3%"><img src="../images/signatures/<% = signature %>"></td><td width="3%" align="left" valign="top"><% = printName %></td><td width="4%" align="left" valign="top">SWPPP INSPECTIONS, INC.</td></tr>
</table><br><br><%
	END IF %></DIV>
<%	RS1.MoveNext 
LOOP
RS1.Close
SET RS1=nothing %>
</body>
</html><%
connSWPPP.Close
Set connSWPPP = Nothing %>