<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") then Response.Redirect("../default.asp") End If
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

inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " &_
	"projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " &_
	"compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, " &_
	"contactFax, contactEmail, reportType, inches, bmpsInPlace, " &_
	"sediment, narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems, horton, hortonSignV, hortonSignLD, vscr, ldscr, forestar" &_
	" FROM Inspections, Projects, Users" &_
	" WHERE inspecID = " & inspecID &_
	" AND Inspections.projectID = Projects.projectID" &_
	" AND Inspections.userID = Users.userID"
'--Response.Write("Inspec: "& inspecSQLSELECT &"<br>")
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
strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
strBody=strBody &"<center><img src='http://www.swpppinspections.com/images/b&wlogoforreport.jpg' width='300'><br><br>"
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
    strBody=strBody &"<tr><td align='right'><b>"
	If rsInspec("forestar") Then
		strBody=strBody &"TPDES Permit #:"
	Else
		strBody=strBody &"On-Site Contact"
	End If
	strBody=strBody &"</b></td><td colspan='3'>"&  Trim(rsInspec("emergencyPhone")) &"</td></tr>"
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
coordSQLSELECT = "SELECT correctiveMods, coordinates, existingBMP FROM Coordinates WHERE inspecID = "& inspecID &" ORDER BY orderby"
Set rsCoord = connSWPPP.Execute(coordSQLSELECT)
If rsInspec("projectState") = "OK" Then
    strBody=strBody &"<p><center><i>Inspectors familiar with the OPDES Permit OKR10 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized, areas used for storage of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles enter and exit the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i></center>"
Else
    strBody=strBody &"<p><center><i>Inspectors familiar with the TPDES Permit TXR150000 and the SWPPP should inspect disturbed areas of the site that have not been finally stabilized, areas used for storage of materials that are exposed to precipitation, structural controls (all erosion and sediment controls), discharge locations, locations where vehicles enter and exit the site, off-site material storage areas, overburden and stockpiles of dirt, borrow areas, equipment staging areas, vehicle repair areas, and fueling areas.</i></center>"
End If
rsInspec.Close
Set rsInspec = Nothing
strBody=strBody &"<p><table border='0' cellpadding='3' width='100%' cellspacing='0'>"
If rsCoord.EOF Then
	strBody=strBody &"<tr><td colspan='2' align='center'><i>There is no coordinate data entered at this time.</i></td></tr>"
Else
	Do While Not rsCoord.EOF
		correctiveMods = Trim(rsCoord("correctiveMods"))
		coordinates = Trim(rsCoord("coordinates"))
		existingBMP = Trim(rsCoord("existingBMP"))
strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>Location (see Site Map):</b></td>	<td width='80%' align='left'>"&  coordinates &"<br></td></tr>"
IF TRIM(rsCoord("existingBMP"))<>"-1" THEN
	strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>Existing BMP:</b></td><td width='80%' align='left'>"&  existingBMP &"<br></td></tr>"
END IF
strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>Corrective Modification:</b></td><td width='80%' align='left'>"&  correctiveMods &"</td></tr>"
strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"
		rsCoord.MoveNext
	Loop
End If ' END No Results Found
rsCoord.Close
Set rsCoord = Nothing
IF narrative <> "" THEN
strBody=strBody &"<TR><TD colspan=4><p><small>"& REPLACE(TRIM(narrative),"#@#","'")&"</small></TD></TR>"
END IF
strBody=strBody &"</table><p><small>"& REPLACE(TRIM(qualifications),"#@#","'")&"</small></p>"
strBody=strBody &"<p><small>I certify under penalty of law that this document and all attachments were prepared under my direction or supervision in accordance with a system designed to assure that qualified personnel properly gathered and evaluated the information submitted. Based on my inquiry of the person or persons who manage the system, or those persons directly responsible for gathering the information, the information is, to the best of my knowledge and belief, true, accurate, and complete. I am aware that there are significant penalties for submitting false information, including the possibility of fine and imprisonment for knowing violations.</small><p><table border='0' cellpadding='2' width='100%' cellspacing='0'>"
strBody=strBody &"<tr><td width='3%' align='left'><b>Name:</b></td><td width='3%' align='left'><b>Print:</b></td><td width='4%' align='left'><b>Inspector:</b></td></tr>"
strBody=strBody &"<tr><td width='3%'><img src='http://www.swpppinspections.com/images/signatures/"&  signature &"'></td><td width='3%' align='left' valign='top'>"&  printName &"</td><td width='4%' align='left' valign='top'>SWPPP INSPECTIONS, INC.</td></tr></table>"
strBody=strBody &"<br><br>"
SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& inspecID
SET RS3=connSWPPP.execute(SQL3)
IF NOT(RS3.EOF) THEN
strBody=strBody &"<div align='center'><a href='http://www.swpppinspections.com/images/sitemap/"& RS3("oImageFileName") &"'>link for Site Map</a></div>"
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
END IF
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



'--------------------- old process mailing original aspmail component - we dont have license--------------------------
'		contentSubject= "Inspection Report for "& TRIM(RS1("projectName")) &" "& TRIM(RS1("projectPhase")) &" on "& TRIM(RS1("inspecDate"))
'		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'		Mailer.FromName    = "Don Wims"
'		Mailer.FromAddress = "dwims@swppp.com"
'		Mailer.RemoteHost = "127.0.0.1"
'		Mailer.Subject    = contentSubject
'		Mailer.BodyText = strBody & strImages & "<Body>"
'		Mailer.ContentType = "text/html"
'--------------------- old process mailing original aspmail component - we dont have license--------------------------



'--------------------- new aspemail component object creation - no license required--------------------------
		contentSubject= "Inspection Report for "& TRIM(RS1("projectName")) &" "& TRIM(RS1("projectPhase")) &" on "& TRIM(RS1("inspecDate"))
		Set Mailer = Server.CreateObject("Persits.Mailsender")
		Mailer.Host = "127.0.0.1"
		Mailer.Port = 25
		Mailer.From = "dwims@swppp.com"
		Mailer.FromName    = "Don Wims"
		Mailer.Subject    = contentSubject
		Mailer.Body = strBody & strImages & "<Body>"
		Mailer.IsHTML = True
'		Mailer.ContentType = "text/html"
'--------------------- new aspemail component object creation - no license required--------------------------



'-- build the recipients list ------------------------------------------------
		DO WHILE NOT RS1.EOF
		    curRights = Trim(RS1("rights"))
            if curRights = "email" then
			    Mailer.AddRecipient Trim(RS1("email")), Trim(RS1("fullName"))
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
If Err <> 0 Then
	Response.Write "An error occurred: " & Err.Description
Else %>
<font color="#008000">Mail Sent Successfully</font>
<%
End If


'--	now it is time to set the released bit on the inspection -------------------------------
		SQL2="UPDATE Inspections SET released=1 WHERE inspecID="& inspecID
		connSWPPP.execute(SQL2)
'--		Response.Write(strBody)
	NEXT
'	Response.End
ELSE
SQL0 = "SELECT i.projectName, i.projectPhase, i.inspecDate, i.inspecID, pu.projectID, i.ReportType, i.released" &_
	" FROM ProjectsUsers pu JOIN Inspections i ON pu.projectID=i.projectID" &_
	" WHERE pu.rights='inspector' AND i.released=0 AND i.userID=pu.userID AND pu.userID="& Session("userID")
If Session("userID") = 42 Then
    SQL0 = SQL0 & " and datediff(m, i.inspecdate, getdate()) <4"
End	If
SQL0 = SQL0 & " ORDER BY i.projectName, i.projectPhase, i.inspecDate DESC"
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
<h1>Send Reports via Email</h1>
<FORM action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
<div align="center">
<table border="0" cellpadding=1 cellspacing=1>
	<tr><th>Project Name|Phase</th><th>Report Date</th><th>Report Type</th><th>send email</th></tr>
<% 	DO WHILE NOT RS0.EOF %>
	<tr><td align="left"><%= RS0("projectName")%>&nbsp;<%= RS0("projectPhase") %></td>
		<td align="left"><%= RS0("inspecDate") %></td>
		<td align="left"><%= RS0("ReportType") %></td>
		<td align="center"><INPUT type="checkbox" name="<%= RS0("projectID")%>:<%= RS0("inspecID")%>" value="<%= RS0("inspecID")%>"></td></tr>
<% 		RS0.MoveNext
	LOOP
RS0.Close
SET RS0=nothing %>
</table></div>

<div align="center"><br><br>To Send These Reports via Email to all Users assigned<br>
	to Receive them and release this report <nobr>click..<input type="submit" value="Send Emails"></nobr></div>
</FORM>
</BODY>
</HTML><%
END IF
connSWPPP.close
SET connSWPPP=nothing %>