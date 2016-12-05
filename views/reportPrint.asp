<%@ Language="VBScript" %>
<!-- #include file="../admin/connSWPPP.asp" --><%
inspecID = Request("inspecID")
inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " & _
	"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " & _
	"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, " & _
	"contactFax, contactEmail, reportType, inches, bmpsInPlace, " & _
	"sediment, narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems" & _
	" FROM Inspections, Projects, Users" & _
	" WHERE inspecID = " & inspecID & _
	" AND Inspections.projectID = Projects.projectID" & _
	" AND Inspections.userID = Users.userID"
'Response.Write(inspecSQLSELECT)
Set rsInspec = connSWPPP.Execute(inspecSQLSELECT)
'Response.Write("signature = " & Trim(rsInspec("signature")) & "<br>")

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
IF IsNull(qualifications) THEN qualifications="" END IF %>
<html>
<head>
<title>SWPPP INSPECTIONS - Print Report</title>
<link rel="stylesheet" type="text/css" href="../global.css">
<style>
	.red{color: #F52006;}
	.orange{color: #F58C06;}
	.yellow{color: #FFC300;}
	.black{color: black;}
</style>
</head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center><img src="../images/color_logo_report.jpg" width="300"><br><br>
<font size="+1"><b>Inspection Report</b></font><hr noshade size="1" width="90%"></center>
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<!-- date -->
	<tr> 
		<td align="right"><b>Date:</b></td>
		<td colspan="3"><% = Trim(rsInspec("inspecDate")) %></td>
	</tr>
	<!-- project name -->
	<tr> 
		<td align="right"><b>Project Name:</b></td>
		<td colspan="3"><% = Trim(rsInspec("projectName")) %>&nbsp;<% = Trim(rsInspec("projectPhase")) %></td>
	</tr>
	<!-- project location -->
	<tr> 
		<td align="right" valign="top"><b>Project Location:</b></td>
		<td colspan="3" valign="top"><% = Trim(rsInspec("projectAddr")) %></td>
	</tr>
	<tr>
		<td align="right">&nbsp;</td>
		<td colspan="3"><% = (Trim(rsInspec("projectCity")) & ", " & rsInspec("projectState") & " " & Trim(rsInspec("projectZip"))) %></td>
	</tr>
	<tr> 
		<td align="right"><b>County:</b></td>
		<td colspan="3"><% = Trim(rsInspec("projectCounty")) %></td>
	</tr>
	<!-- On-Site Contact -->
	<tr> 
		<td align="right"><b>On-Site Contact:</b></td>
		<td colspan="3"><% = Trim(rsInspec("onsiteContact")) %></td>
	</tr>
	<!-- office phone number --><% 
	If Len(Trim(rsInspec("officePhone"))) > 0 then %>
	<tr> 
		<td align="right"><b>On-Site Contact:</b></td>
		<td colspan="3"><% = Trim(rsInspec("officePhone")) %></td>
	</tr><%
	End If %>
	<!-- emergency phone number --><% 
	If Len(Trim(rsInspec("emergencyPhone"))) > 0 then %>
	<tr> 
		<td align="right"><b>On-Site Contact:</b></td>
		<td colspan="3"><% = Trim(rsInspec("emergencyPhone")) %></td>
	</tr><%
	End If %>
	<!-- company, contact -->
	<tr> 
		<td align="right"><b>Company:</b></td>
		<td><% = Trim(rsInspec("compName")) %></td>
		<td align="right"><b>Main Office Contact:</b></td>
		<td><% = Trim(rsInspec("compContact")) %></td>
	</tr>
	<!-- address 1, phone -->
	<tr> 
		<td align="right" valign="top"><b>Address:</b></td>
		<td><% = Trim(rsInspec("compAddr")) %> <% If Trim(rsInspec("compAddr2")) <> "" Then Response.Write("<br>" & Trim(rsInspec("compAddr2"))) End If %></td>
		<td align="right"><b>Phone:</b></td>
		<td><% = Trim(rsInspec("contactPhone")) %></td>
	</tr>
	<!-- address 2, fax -->
	<tr> 
		<td align="right"><b>&nbsp;</b></td>
		<td><% = (Trim(rsInspec("compCity")) & ", " & rsInspec("compState") & " " & Trim(rsInspec("compZip"))) %></td>
		<td align="right"><b>Fax:</b></td>
		<td><% = Trim(rsInspec("contactFax")) %></td>
	</tr>
	<!-- main telephone -->
	<tr> 
		<td align="right"><b>Main Telephone Number:</b></td>
		<td><% = Trim(rsInspec("compPhone")) %></td>
		<td align="right"><b>E-Mail:</b></td>
		<td><% = Trim(rsInspec("contactEmail")) %></td>
	</tr>
	<!-- type of report, inches of rain -->
	<tr> 
		<td align="right"><b>Type of Report:</b></td>
		<td><% = reportType %></td>
<%  IF inches>-1 THEN %>
		<td align="right"><b>Inches of Rain:</b></td>
		<td><% If reportType <> "biWeekly" Then Response.Write(inches) Else Response.Write("N/A") %></td>
<%	ELSE %><td></td>
<% END IF %>
	</tr>
	<tr> 
<%  IF rsInspec("bmpsInPlace")>-1 THEN %>
		<td align="right"><b>Are BMPs in place?</b></td>
		<td><% = bmpsInPlace %></td>
<%  END IF
	IF rsInspec("sediment")>-1 THEN %>
		<td align="right"><b>Sediment Loss or Pollution?</b></td>
		<td><% = sediment %></td>
<% 	END IF %>
	</tr>
</table><%
signature = Trim(rsInspec("signature"))

coordSQLSELECT = "SELECT coID, coordinates, existingBMP, correctiveMods, orderby, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly" &_
	" FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)%>
<p>
	<center><div style="font-size: 10px"><%
If rsInspec("projectState") = "OK" Then %>
<i>Inspectors familiar with the OPDES Permit OKR10 and the SWPPP should inspect 
disturbed areas of the site that have not been finally stabilized, areas used for 
storage of materials that are exposed to precipitation, structural controls (all 
erosion and sediment controls), discharge locations, locations where vehicles enter 
and exit the site, off-site material storage areas, overburden and stockpiles of dirt, 
borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i> 
<% Else %>
<i>Inspectors familiar with the TPDES Permit TXR150000 and the SWPPP should inspect 
disturbed areas of the site that have not been finally stabilized, areas used for 
storage of materials that are exposed to precipitation, structural controls (all 
erosion and sediment controls), discharge locations, locations where vehicles enter 
and exit the site, off-site material storage areas, overburden and stockpiles of dirt, 
borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i>
<% End If %>
</div></center>
<p> 
<table border="0" cellpadding="3" width="100%" cellspacing="0"><%
If rsInspec("compliance") Then
	Response.Write("<tr><td colspan='2' align='center'><h2>SITE IS IN COMPLIANCE WITH THE SWPPP</h2></td></tr>")
Else 
    If rsCoord.EOF Then
	    Response.Write("<tr><td colspan='2' align='center'><i>There is no " & _
		    "coordinate data entered at this time.</i></td></tr>")
    Else
        applyScoring = False
	    'if rsInspec("includeItems")=True & Session("seeScoring")=True Then applyScoring = True End If
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
            If infoOnly = True Then %>
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
		        <% END IF %>
		        <tr valign='top'><td width='20%' align='right'><b>action needed:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=correctiveMods%></td></tr>
		        <% IF applyScoring and repeat THEN %>
			        <tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class = '<%=scoring_class%>'><%=age%><br></td></tr>
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
If rsInspec("projectState") = "OK" Then %>
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
<p><div style="font-size: 10px"><%= REPLACE(TRIM(qualifications),"#@#","'")%></div></p>
<p><div style="font-size: 10px">I certify under penalty of law that this document and all attachments 
	were prepared under my direction or supervision in accordance with a system 
	designed to assure that qualified personnel properly gathered and evaluated 
	the information submitted. Based on my inquiry of the person or persons who 
	manage the system, or those persons directly responsible for gathering the 
	information, the information is, to the best of my knowledge and belief, true, 
	accurate, and complete. I am aware that there are significant penalties for 
	submitting false information, including the possibility of fine and imprisonment 
	for knowing violations.</div> 
<p> 
<table border="0" cellpadding="2" width="100%" cellspacing="0">
	<tr> 
		<td width="3%" align="left"><b>Name:</b></td>
		<td width="3%" align="left"><b>Print:</b></td>
		<td width="4%" align="left"><b>Inspector:</b></td>
	</tr>
	<tr> 
		<td width="3%"><img src="../images/signatures/<% = signature %>"></td>
		<td width="3%" align="left" valign="top"><% = printName %></td>
		<td width="4%" align="left" valign="top">SWPPP INSPECTIONS, INC.</td>
	</tr>
</table>
<% 
rsCoord.Close
Set rsCoord = Nothing

rsInspec.Close 
Set rsInspec = Nothing

connSWPPP.Close
Set connSWPPP = Nothing

%>
<br>
<br>
</body>
</html>