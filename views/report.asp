<%@ Language="VBScript" %>
<% If _
	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") And _
	Not Session("validErosion") _
Then
'	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
	Session("adminReturnTo") = Request.ServerVariables("path_translated") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If

inspecID = Request("inspecID")
SQL0= "SELECT projectName, projectPhase, inspecDate FROM Inspections WHERE inspecID="& inspecID
%><!-- #include virtual="admin/connSWPPP.asp" --><%
SET RS0=connSWPPP.execute(SQL0)
projectName=TRIM(RS0("projectName"))
projectPhase=TRIM(RS0("projectPhase"))
inspecDate=RS0("inspecDate") %>
<html>
<head>
	<title>SWPPP INSPECTIONS : <% = projectName %> <%= projectPhase%> on <% = inspecDate %></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../global.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="../js/popWindow.js"></script>
</head>
<body onUnload="closePopWin()";>
	<!-- #include virtual="header.inc" -->
	<h1>Inspection for <% = projectName %>&nbsp;<%= projectPhase%> on <% = inspecDate %></h1>
	<div class="four columns alpha">
		<h3>Available Reports</h3>
		<div class="side-link-small">
			<a href="reportPrint.asp?inspecID=<% = inspecID %>" target="_blank">Report</a>
		</div>
		<% SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=-1 ORDER BY oitSortByVal asc"
		SET RS1=connSWPPP.execute(SQL1)
		DO WHILE NOT RS1.EOF 
			dirName=Trim(RS1("oitName"))
			fileDesc= TRIM(RS1("oitDesc"))
			SQLa="sp_oImagesByType "& inspecID &",'"& RS1("oitID") &"'"
			SET RSa=connSWPPP.Execute(SQLa)
			cnt1=1
			curOITDesc=""
			DO WHILE NOT(RSa.EOF)
				thisFileDesc=fileDesc
				if curOITDesc=fileDesc then
					cnt1=cnt1+1
				else
					cnt1=1
					curOITDesc=fileDesc
				end if
				if (cnt1>1) then 
					thisFileDesc=thisFileDesc &" "& cnt1 
				end if
				IF 	dirName = "sitemap" THEN %>
					<div class="side-link-small">
						<a href="<% = "../images/"& dirName &"/"& RSa("oImageFileName") %>" target="_blank"><%= thisFileDesc%></a>
					</div>
				<% ELSE
					If Not Session("validErosion") Then %>
						<div class="side-link-small">
							<a href="<% = "../images/"& dirName &"/"& RSa("oImageFileName") %>" target="_blank"><%= thisFileDesc%></a>
						</div>
					<% End If
				END IF
			RSa.MoveNext
			LOOP 
		RS1.MoveNext
		LOOP %>
	</div>
	<!----------------------------------- Images ------------------------------>
	<div class="eight columns omega">
		<h3>Site Images:</h3>
		<% IF NOT Session("noImages") THEN
			imgSQLSELECT = "SELECT imageID, largeImage, smallImage, description FROM Images WHERE inspecID = " & inspecID
			Set rsImages = connSWPPP.execute(imgSQLSELECT)

			If Not rsImages.EOF Then %>
				
				<% Do While Not rsImages.EOF
					iDataRows = iDataRows + 1
					If iDataRows > 3 Then
						Response.Write("</tr>" & VBCrLf & "<tr>")
						iDataRows = 1
					End If %>
					<a href="<%= "../images/lg/" & Trim(rsImages("largeImage"))%> target="_blank"><%= Trim(rsImages("description")) %><br>
					<% If Right(Trim(rsImages("smallImage")),3)="pdf" then %>
						<img src="../images/acrobat.gif" width="87" height="30" border="0" alt="Acrobat PDF Doc">
					<% else %>
						<img src="<%= "../images/sm/" & Trim(rsImages("smallImage")) %>" border="0" alt="<%= Trim(rsImages("smallImage")) %>">
					<% end if %>
					</a><br/>
				<% rsImages.MoveNext
				Loop %>
			<% END IF	'--- noImages Check
		rsImages.Close
		Set rsImages = Nothing
		End If
	connSWPPP.Close
	Set connSWPPP = Nothing %>
	</div>
</body>
</html>