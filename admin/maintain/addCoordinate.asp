<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") &_
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

inspecID = Session("inspecID")

If Request.Form.Count > 0 Then
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
%>
<!-- #include virtual="admin/connSWPPP.asp" -->
<%	coordSQLINSERT = "INSERT INTO Coordinates (" & _
		"inspecID, coordinates, existingBMP, correctiveMods, orderby" & _
		") VALUES (" & _
		inspecID & _
		", '" & strQuoteReplace(Request("coordinates")) & "'" & _
		", -1" & _
		", '" & strQuoteReplace(Request("correctiveMods")) & "'" & _
		", " & Request("orderby") & _
		")"
	' Response.Write(coordSQLINSERT & "<br><br>")
	connSWPPP.Execute(coordSQLINSERT)
	
	connSWPPP.Close
	Set connSWPPP = Nothing	
End If %>
<html>
<head>
<title>SWPPP INSPECTIONS : Add Locations</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
<script language="JavaScript" src="../js/validCoordinates.js"></script>
<script language="JavaScript" src="../js/validCoordinates1.2.js"></script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Add Location</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
	<form method="post" action="addCoordinate.asp" onSubmit="return isReady(this);">
		<input type="hidden" name="inspecID" value="<%= inspecID %>">
		<tr><td colspan="2" align="center"> <input type="button" value="View Reports" onClick="location='viewReports.asp'; return false";></td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<!-- coordinates -->
		<tr><td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr><td align="right" bgcolor="#eeeeee"><b>Locations:</b></td>
			<td bgcolor="#999999"><input name="coordinates" type="text" value="" size="50" maxlength="50"></td></tr>
		<!-- bmp -->
<!--		<INPUT type="hidden" name="existingBMP" value="-1">-->
		<!-- corrective mods -->
		<tr><td align="right" valign="top" bgcolor="#eeeeee" nowrap><b>Corrective Mods:</b></td>
			<td bgcolor="#999999"><textarea name="correctiveMods" cols="50" rows="5"></textarea></td></tr>
		<tr><td align="right" nowrap bgcolor="#eeeeee"><b>List Order:</b></td>
			<td bgcolor="#999999"><input type="text" name="orderby" size="4" maxlength="4"></td></tr>
		<tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<tr><td colspan="2"><small>I certify under penalty of law that this document 
				and all attachments were prepared under my direction or supervision 
				in accordance with a system designed to assure that qualified 
				personnel properly gathered and evaluated the information submitted. 
				Based on my inquiry of the person or persons who manage the system, 
				or those persons directly responsible for gathering the information, 
				the information is, to the best of my knowledge and belief, true, 
				accurate, and complete. I am aware that there are significant 
				penalties for submitting false information, including the possibility 
				of fine and imprisonment for knowing violations. </small></td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<tr><td colspan="2" align="center"> <input type="submit" value="Submit Coordinate"> 
			<INPUT type="button" value="Return to Report View" onClick="window.navigate('editReport.asp');"></td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
	</form>
</table>
</body>
</html>