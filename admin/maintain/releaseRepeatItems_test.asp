<%Response.Buffer = False%>
<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and not Session("validAdmin") then Response.Redirect("../default.asp") End If
%><!-- #INCLUDE FILE="../connSWPPP.asp" --><%

Server.ScriptTimeout=1500
'Response.Write(Request.Form.Count & "<br>")
inspecID = Request("inspecID")

'--	Item is ProjectsUsers.projectID &":"& Inspections.inspecID -------------------
'--	Request(Item) is Inspections.inspecID ----------------------------------------
'-- need to create the Email content ---------------------------------------------
strBody=""

inspecSQLSELECT = "SELECT inspecDate, Inspections.projectName, Inspections.projectPhase, projectAddr, projectCity, projectState, " &_
    "projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, compName, " &_
    "compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax, " &_
    "contactEmail, reportType, inches, bmpsInPlace, sediment, " &_
    "narrative, firstName, lastName, signature, qualifications, includeItems, compliance, totalItems, completedItems, horton, sentRepeatItemReport" &_
        " FROM Inspections, Projects, Users" &_
        " WHERE inspecID = " & inspecID &_
        " AND Inspections.projectID = Projects.projectID" &_
        " AND Inspections.userID = Users.userID"
'Response.Write("Inspec: "& inspecSQLSELECT &"<br>")
Set rsInspec = connSWPPP.Execute(inspecSQLSELECT)
printName = Trim(rsInspec("firstName")) & " " & Trim(rsInspec("lastName"))

projectName = Trim(rsInspec("projectName"))
projectPhase = Trim(rsInspec("projectPhase"))
inspecDate = rsInspec("inspecDate")

strBody=strBody &"<head><style>"
strBody=strBody &".red{color: #F52006;}"
strBody=strBody &".green{color: green;}"
strBody=strBody &".black{color: black;}"
strBody=strBody &".ld{font-weight: bold;}"
strBody=strBody &"</style></head>"
strBody=strBody &"<body bgcolor='#ffffff' marginwidth='30' leftmargin='30' marginheight='15' topmargin='15'>"
strBody=strBody &"<center><img src='http://www.swpppinspections.com/images/color_logo_report.jpg' width='300'><br><br>"
strBody=strBody &"<font size='+1'><b>Repeat Item Report</b></font><br/>"
strBody=strBody &"<font size='+1'><b>" & projectName & " " & projectPhase & "</b></font><br/>"
strBody=strBody &"<font size='+1'><b>" & inspecDate & "</center><br/>"

coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)

strBody=strBody &"<h3>Repeat Items</h3>"
strBody=strBody &"<p><table border='0' cellpadding='3' width='100%' cellspacing='0'>"
strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"
If rsCoord.EOF Then
    strBody=strBody &"<h5>No Items Found</h5>"
Else
    applyScoring = rsInspec("includeItems")
    currentDate = date()
    send_email = False
    Do While Not rsCoord.EOF
        coID = rsCoord("coID")
        correctiveMods = Trim(rsCoord("correctiveMods"))
        orderby = rsCoord("orderby")
        coordinates = Trim(rsCoord("coordinates"))
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
        dis = rsCoord("discharge")
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
        If applyScoring Then
            If assignDate = "" Then
                age = 0
            Else
                age = datediff("d",assignDate,currentDate) 
            End If
        If age > 6 Then
            scoring_class = "red"
        End If
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
        If dis = True Then
        	correctiveMods = "(discharge) " & correctiveMods
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
        If infoOnly = True or NLN = True Then
            do_nothing = 1 
        Elseif  repeat and age > 0 THEN
            send_email = True
            scoring_class = "red"
            If useAddress Then
                    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class='red'>"&  locationName &"<br></td></tr>"
                    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>address:</b></td>	<td width='80%' align='left' class='red'>"&  address &"<br></td></tr>"
                Else
                    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>location:</b></td>	<td width='80%' align='left' class='red'>"&  coordinates &"<br></td></tr>"
                End If
                strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>action needed:</b></td><td width='80%' align='left' class='red'>"&  correctiveMods &"</td></tr>"
                If applyScoring Then
                    strBody=strBody &"<tr valign='top'><td width='20%' align='right'><b>item age:</b></td><td width='80%' align='left' class='red'>"&  age &" days<br></td></tr>"
                End If
                strBody=strBody &"<tr><td colspan='2'><hr noshade size='1' align='center' width='90%'></td></tr>"  & vbCrLf
            End If
        End If 'end applyScoring
        rsCoord.MoveNext
    Loop
End If ' END No Results Found

strBody=strBody &"</table><br><center>Complete Report: <a href='http://www.swppp.com/views/reportPrint.asp?inspecID="& inspecID &"'>http://www.swppp.com/views/reportPrint.asp?inspecID="& inspecID &"</a>"
SQL3="SELECT oImageFileName FROM OptionalImages WHERE oitID=12 AND inspecID="& inspecID
SET RS3=connSWPPP.execute(SQL3)
IF NOT(RS3.EOF) THEN
    strBody=strBody &"<br>Site Map: <a href='http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName")) &"'>http://www.swpppinspections.com/images/sitemap/"& TRIM(RS3("oImageFileName")) &"</a>"
END IF
strBody=strBody &"<br>Website: <a href='http://www.swppp.com'>www.swppp.com</a></center></Body>"

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
Set Mailer = Server.CreateObject("Persits.MailSender")
Mailer.FromName    = "SWPPP.COM"
Mailer.From = "dwims@swppp.com"
Mailer.Host = "127.0.0.1"
Mailer.Subject    = contentSubject
Mailer.Body = strBody & strImages & "<Body>"
BodyText = strBody & strImages & "<Body>"
Mailer.isHTML     = True
Mailer.AddAddress "bradleyclare@gmail.com", "Brad Leishman" %>
            
            
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