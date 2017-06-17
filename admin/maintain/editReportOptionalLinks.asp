<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID = Session("inspecID")
IF Request("inspecID")<>"" THEN 
	inspecID = Request("inspecID") 
	Session("inspecID")=inspecID
END IF
%><!-- #include file="../connSWPPP.asp" --><%
If Request.Form.Count > 0 Then	
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	
	if Request.Form("submit") = "Return to Report w/o Saving" then
		Response.Redirect("editReport.asp?inspecID=" + inspecID)
	End If
	userID = Session("userID")
	'--	Creat Optional Image Associations -------------------------------
	SQLa="DELETE FROM OptionalImages WHERE inspecID="& inspecID 
	connSWPPP.Execute(SQLa)
	SQLa="SELECT * FROM OptionalImagesTypes ORDER BY oitSortByVal asc"
	SET RSa=connSWPPP.Execute(SQLa)
	DO WHILE NOT RSa.EOF
		cnt=1
'Response.Write(Request(Trim(RSa("oitName"))) & "<BR/>")
		tList=SPLIT(Request(Trim(RSa("oitName"))),":")
		For m=1 to Ubound(tList) Step 1
'--	dbo.sp_AddOptImage 	@oImageName char(50),@oIamgeDesc char(50),@inspecID int,@oitID int,@oImageFileName char(50),@oOrder smallint
			SQLb="EXEC sp_AddOptImage null,null,"& inspecID &","& RSa("oitID") &",'"& Trim(tList(m-1)) &"',"& cnt
         'Response.Write(SQLb & "<br/>")
         connSWPPP.execute(SQLb)
			cnt=cnt+1
		Next
		RSa.MoveNext
	LOOP 
   'if Len(SQLb) > 1 then connSWPPP.execute(SQLb) End If
		
	if Request.Form("submit") = "Save Updates and Return to Report" then
		'Response.Redirect("editReport.asp?inspecID=" + inspecID)
	End If
		
End If
inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, projectAddr, projectCity, projectState" & _
	", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, i.projectID, compName" & _
	", compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax" & _
	", contactEmail, reportType, inches, bmpsInPlace, sediment, userID" & _
	" FROM Inspections as i, Projects as p" & _
	" WHERE i.projectID = p.projectID AND inspecID = " & inspecID
'--Rsponse.Write(inspecSQLSELECT & "<br>")
Set rsReport = connSWPPP.execute(inspecSQLSELECT)
baseDir = "D:\Inetpub\wwwroot\SWPPP\"
%>

<html>
<head>
	<title>SWPPP INSPECTIONS : Edit Optional Links Report</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<STYLE>
	select.long	{ font-size:xx-small;	}
	</STYLE>
	<script type="text/javascript" language="JavaScript" src="../js/validReports.js"></script>
	<script type="text/javascript" language="JavaScript" src="../js/validReports1.2.js"></script>
	<script type="text/javascript" language="JavaScript1.2"><!--
// we Can't just use the same transfer function for both directions because
// the hidden input keys off of the t2 value solely...-->
function addOption(t1, t2, t3) {
    var index = t3.selectedIndex;
    if (index > -1) {
        var newoption = new Option(t3.options[index].text, t3.options[index].value, true, true);
        t2.options[t2.length] = newoption;
        if (!document.getElementById) history.go(0);
        t3.options[index] = null;
        t3.selectedIndex = 0;
		var tempStr="";
		for(var i=0; i<(t2.length) ;i++){
			tempStr=tempStr + (t2.options[i].value) + ":" ;
		}
		t1.value=tempStr;
    }
}
function delOption(t1, t3, t2) {
    var index = t3.selectedIndex;
    if (index > -1) {
        var newoption = new Option(t3.options[index].text, t3.options[index].value, true, true);
        t2.options[t2.length] = newoption;
        if (!document.getElementById) history.go(0);
        t3.options[index] = null;
        t3.selectedIndex = 0;
		var tempStr="";
		for(var i=0; i<(t3.length) ;i++){
			tempStr=tempStr + (t3.options[i].value) + ":" ;
		}
		t1.value=tempStr;
    }
}
function swapOption(t1, t2, slideDir) {
	var curIndex = t2.selectedIndex;
	var swapIndex= curIndex;
	var maxIndex= t2.length;
	if (curIndex > -1) {
		(slideDir=="up") ? (swapIndex=curIndex-1):(swapIndex=curIndex+1);
		if ((swapIndex>-1) && (swapIndex<t2.length)) {
			var newOption = new Option(t2.options[swapIndex].text, t2.options[swapIndex].value, true, true);
			t2.options[maxIndex] = newOption;
			t2.options[swapIndex].text=t2.options[curIndex].text;
			t2.options[swapIndex].value=t2.options[curIndex].value;
			t2.options[curIndex].text=t2.options[maxIndex].text;
			t2.options[curIndex].value=t2.options[maxIndex].value;
			t2.options[maxIndex] = null;
			t2.selectedIndex=swapIndex;
			var tempStr="";
			for(var i=0; i<(t2.length) ;i++){
				tempStr=tempStr + (t2.options[i].value) + ":" ;
			}
			t1.value=tempStr;
		}
	}	
}
</script>
</head>
<body>

<!-- #include file="../adminHeader2.inc" -->
<h1>Optional Project Links</h1>
<h2>Edit Optional Links to Report for <% = Trim(rsReport("projectName")) %></h2>	
<form id="theForm" method="post" action="<% = Request.ServerVariables("script_name") %>" onsubmit="return isReady(this)";>
	<input type="hidden" name="inspecID" value="<% = inspecID %>"/>
	<input type="hidden" name="projectID" value="<% = rsReport("projectID") %>"/>
<!-- ------------- Optional Links ----------------------------------------------------- -->
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">
		<tr><td colspan=3 align="center"><input type="submit" name="submit" value="Return to Report w/o Saving" /></td>
		<td colspan=3 align="center"><input name="submit" type="submit" value="Save Updates and Return to Report"/></td></tr>
		<tr><td colspan="5"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999" colspan=5><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr valign="bottom">
		<td align="right" bgcolor="#eeeeee" style="border-bottom: 1px solid black;"><font size="-2"><strong>Type</strong></font></td>
		<td bgcolor="#999999" style="padding: 0cm; margin: 0cm; border-bottom: 1px solid black;">&nbsp;</td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black; padding: 0cm; margin: 0cm;"><font size="-2"><strong>Current Links</strong></font></td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black;" align="center"><font size="-2"><strong><nobr>rem | add</nobr></strong></font></td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black;"><font size="-2"><strong>Available Links</strong></font></td>
		<td bgcolor="#999999" style="border-bottom: 1px solid black;">&nbsp;</td></tr>
<%	
SQL1="SELECT * FROM OptionalImagesTypes WHERE oitSortByVal >=0 ORDER BY oitSortByVal asc"
SET RS1=connSWPPP.execute(SQL1)
DO WHILE NOT RS1.EOF
	oitID=RS1("oitID")
	dirName=Trim(RS1("oitName"))
	oitDesc=Trim(RS1("oitDesc"))
	t1=dirName
	t2=dirName &"DN"
	t3=dirName &"UP"

	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objTemp = FSO.GetFolder(baseDir & "images\"& dirName &"\")
	Set TempImage = objTemp.Files 
	SQLa="sp_oImagesByType "& inspecID &",'"& RS1("oitID") &"'" 
	SET RSa=connSWPPP.Execute(SQLa)
	tempStrOfFileNames="" %>
		<tr valign="top">
			<td align="right" bgcolor="#eeeeee"><font size="-1"><strong><%= oitDesc%>:</strong></font></td>
			<td width='0' bgcolor="#999999" style="padding: 0cm; margin: 0cm;" align="right">
				<BUTTON type="button" onClick="swapOption(<%= t1%>, <%= t2%>, 'up');">&uarr;</BUTTON><br>
				<BUTTON type="button" onClick="swapOption(<%= t1%>, <%= t2%>, 'dn');">&darr;</BUTTON></td>	
			<td bgcolor="#999999" nowrap align="left" style="padding: 0cm; margin: 0cm;" align=left>
				<select name="<%=dirName%>DN" size="3" class="long" style="padding: 0cm; margin: 0cm;" align=left>
<% 	DO WHILE NOT(RSa.EOF) %><OPTION value="<%= Trim(RSa("oImageFileName"))%>"><%= Trim(RSa("oImageFileName"))%></OPTION>
<%		tempStrOfFileNames=tempStrOfFileNames & TRIM(RSa("oImageFileName"))&":"
		RSa.MoveNext
	LOOP %>		</SELECT></td>				
			<td bgcolor="#999999" align="center">
				<BUTTON type="button" onClick="delOption(<%= t1%>, <%= t2%>, <%= t3%>);">--&gt;</BUTTON>
				<BUTTON type="button" onClick="addOption(<%= t1%>, <%= t2%>, <%= t3%>);">&lt;--</BUTTON></td>
			<td bgcolor="#999999"><select name="<%=dirName%>UP"  class="long">
<%	For Each Item In TempImage
		shortName = Item.Name 
		IF InStr(tempStrOfFileNames, shortName)=0 THEN %><option value="<% = Trim(shortName) %>"><% = Trim(shortName) %></option>
<%		End If  '-- this weeds out those already selected fileNames
	Next 
	Set objTemp = Nothing
	Set TempImage = Nothing %>
				</select></td>
			<td bgcolor="#999999" nowrap align="right"><input type="button" style="width:150;"
				value="Upload File/Image" onClick="location='upImageRpt.asp?inspecID=<%= inspecID %>&oitID=<%= oitID %>'; return false";>
				<input type="hidden" name="<%= t1%>" value="<%= tempStrOfFileNames%>"></td></tr>
<% 	RS1.MoveNext
LOOP %>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"/></td>
			<td bgcolor="#999999" colspan=5><img src="../../images/dot.gif" width="5" height="5"/></td>
		</tr><tr><td colspan="5"><img src="../../images/dot.gif" width="5" height="5"/></td></tr>
		<tr><td colspan="6">&nbsp;</td></tr>
		<tr><td colspan=3 align="center"><input type="submit" name="submit" value="Return to Report w/o Saving" /></td>
		<td colspan=3 align="center"><input name="submit" type="submit" value="Save Updates and Return to Report"/></td></tr>
</table>

<%
connSWPPP.Close 
Set connSWPPP = Nothing 
%>
<br><br>
</form>
</body>
</html>