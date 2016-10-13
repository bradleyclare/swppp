<%@ Language="VBScript" %>
<%
If _
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
SQL0= "SELECT projectID, projectName, projectPhase, inspecDate, completedItems, totalItems, includeItems, compliance FROM Inspections WHERE inspecID="& inspecID
%><!-- #include file="../admin/connSWPPP.asp" --><%
SET RS0=connSWPPP.execute(SQL0)
projectID      = RS0("projectID")
projectName    = TRIM(RS0("projectName"))
projectPhase   = TRIM(RS0("projectPhase"))
inspecDate     = RS0("inspecDate") 
totalItems     = RS0("totalItems")
completedItems = RS0("completedItems") 
includeItems   = RS0("includeItems") 
compliance     = RS0("compliance") 
openItems      = totalItems - completedItems %>
<html>
<head>
	<title>SWPPP INSPECTIONS : <% = projectName %>&nbsp;<%= projectPhase%> on <% = inspecDate %></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../global.css" rel="stylesheet" type="text/css">
	<script language="JavaScript" src="../js/popWindow.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" onUnload="closePopWin()";>
<!-- #include file="../header2.inc" -->
<h1>Inspection for <% = projectName %>&nbsp;<%= projectPhase%> on <% = inspecDate %></h1>
      <p class="indent30"><a href="reportPrint.asp?inspecID=<% = inspecID %>" target="_blank">Report</a><%
SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal>=-1 ORDER BY oitSortByVal asc"
SET RS1=connSWPPP.execute(SQL1)
DO WHILE NOT RS1.EOF %><p class="indent30"><%
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
		if (cnt1>1) then thisFileDesc=thisFileDesc &" "& cnt1 end if
		IF 	dirName = "sitemap" THEN %>
      	<a href="<% = "../images/"& dirName &"/"& RSa("oImageFileName") %>" target="_blank"><%= thisFileDesc%></a><br>
<%		ELSE
			If Not Session("validErosion") Then %>
      	<a href="<% = "../images/"& dirName &"/"& RSa("oImageFileName") %>" target="_blank"><%= thisFileDesc%></a><br>
<%			End If
		END IF
		RSa.MoveNext
	LOOP
	RS1.MoveNext
LOOP %>
<!----------------------------------- Images ------------------------------>
<% IF NOT Session("noImages") THEN
imgSQLSELECT = "SELECT imageID, largeImage, smallImage, description FROM Images WHERE inspecID = " & inspecID
Set rsImages = connSWPPP.execute(imgSQLSELECT)

If Not rsImages.EOF Then %>
<div class="indent30"><b>Site Images:</b><br><br>
<table cellspacing="0" cellpadding="4" width="90%" border="0">
	<tr><%
Do While Not rsImages.EOF
	iDataRows = iDataRows + 1
	If iDataRows > 3 Then
		Response.Write("</tr>" & VBCrLf & "<tr>")
		iDataRows = 1
	End If %>
		<td align="center"><a href="<%= "../images/lg/" & Trim(rsImages("largeImage")) %>"
			target="_blank"><%= Trim(rsImages("description")) %><br>
			<% If Right(Trim(rsImages("smallImage")),3)="pdf" then %>
			<img src="../images/acrobat.gif" width="87" height="30" border="0" alt="Acrobat PDF Doc">
			<% else %>
			<img src="<%= "../images/sm/" & Trim(rsImages("smallImage")) %>" border="0"
				alt="<%= Trim(rsImages("smallImage")) %>">
			<% end if %>
			</a></td>
<%	rsImages.MoveNext
Loop %>
	</tr>
</table><br><br>
</div>
<% END IF	'--- noImages Check
rsImages.Close
Set rsImages = Nothing
End If 
If includeItems Then %>
<a href="openActionItems.asp?pID=<%= projectID%>&inspecID=<% = inspecID %>" target="_blank">(<%=openItems%>) Open Items</a><br/>
<a href="completedActionItems.asp?pID=<%= projectID%>&inspecID=<% = inspecID %>" target="_blank">(<%=completedItems%>) Completed Items</a>
<% End If
connSWPPP.Close
Set connSWPPP = Nothing %>
</td></tr></table>
</body>
</html>