<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") and not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request("query_string")
	Response.Redirect("loginUser.asp")
End If
%><!-- #include virtual="admin/connSWPPP.asp" --><%
If Request.Form.Count > 0 Then
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	
	SQLUPDATE = "UPDATE Coordinates SET" & _
		" coordinates = '" & strQuoteReplace(Replace(Request("coordinates"), "--","–")) & "'" & _
		", existingBMP = '" & strQuoteReplace(Request("existingBMP")) & "'" & _
		", correctiveMods = '" & strQuoteReplace(Replace(Request("correctiveMods"), "--","–")) & "'" & _
		", orderby = " & Request("orderby") & _
		" WHERE coID = " & Request("coID")
	' Response.Write(coordSQLINSERT & "<br><br>")
	connSWPPP.Execute(SQLUPDATE)
	
	connSWPPP.Close
	Set connSWPPP = Nothing
	
	Response.Redirect("editReport.asp")
	
Else
	SQLSELECT = "SELECT inspecID, coordinates, existingBMP, correctiveMods, orderby" & _
		" FROM Coordinates" & _
		" WHERE coID = " & Request("coID")		
	Set connCoord = connSWPPP.execute(SQLSELECT)	
End If
inspecID=Session("inspecID") %>
<html>
<head>
<title>SWPPP INSPECTIONS : Add Location</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
<script language="JavaScript" src="../js/validCoordinates.js"></script>
<script language="JavaScript" src="../js/validCoordinates1.2.js"></script>
</head>
<body>
<!-- #include virtual="admin/adminHeader2.inc" -->
<h1>Edit Location</h1>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
	<form method="post" action="<% = Request.ServerVariables("script_name") %>">
		<input type="hidden" name="coID" value="<%= Request("coID") %>">
		<tr> 
			<td colspan="2" align="center"> <input type="button" value="Delete Coordinate" 
			onClick="location='deleteCoordinate.asp?coID=<%= Request("coID") %>'; return false";> 
			</td>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;</td>
		</tr>
		<!-- coordinates -->
		<tr> 
			<td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Location:</b></td>
			<td bgcolor="#999999"><input name="coordinates" type="text" size="50" maxlength="50" 
			value="<% = Trim(connCoord("coordinates")) %>"></td>
		</tr>
		<!-- bmp -->
<% IF TRIM(connCoord("existingBMP")) = "-1" THEN %>
		<INPUT type="hidden" name="existingBMP" value="<%= connCoord("existingBMP")%>">
<% ELSE %>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Existing BMP:</b></td>
			<td bgcolor="#999999"><input name="existingBMP" type="text" size="50" maxlength="50" 
			value="<% = Trim(connCoord("existingBMP")) %>"></td>
		</tr>
<% END IF %>
		<!-- corrective mods -->
		<tr> 
			<td align="right" valign="top" bgcolor="#eeeeee" nowrap><b>Corrective 
				Mods:</b></td>
			<td bgcolor="#999999"><textarea name="correctiveMods" cols="50" 
			rows="5"><% = connCoord("correctiveMods") %></textarea></td>
		</tr>
		<tr> 
			<td align="right" nowrap bgcolor="#eeeeee"><b>List Order:</b></td>
			<td bgcolor="#999999"><input type="text" name="orderby" size="4" maxlength="4" 
			value="<% = connCoord("orderby") %>"></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="2"> <small>I certify under penalty of law that this document 
				and all attachments were prepared under my direction or supervision 
				in accordance with a system designed to assure that qualified 
				personnel properly gathered and evaluated the information submitted. 
				Based on my inquiry of the person or persons who manage the system, 
				or those persons directly responsible for gathering the information, 
				the information is, to the best of my knowledge and belief, true, 
				accurate, and complete. I am aware that there are significant 
				penalties for submitting false information, including the possibility 
				of fine and imprisonment for knowing violations. </small></td>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> <input type="submit" value="Modify Coordinate"></td>
		</tr>
	</form>
</table>
</body>
</html>
<%
connCoord.Close
Set connCoord = Nothing

connSWPPP.Close
Set connSWPPP = Nothing
%>