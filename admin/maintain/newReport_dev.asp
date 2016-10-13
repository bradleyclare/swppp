<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
%>
<!-- #include file="../connSWPPP.asp" -->
<%
If Request.Form.Count > 0 Then
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	
	userID = Session("userID")
	inspector = strQuoteReplace(Request("inspector"))
	
	If inspector <> "" Then userID = inspector End If
	
	companySQLINSERT = "INSERT INTO Companies (" & _
		"companyName" & _
		") VALUES (" & _
		"'" & strQuoteReplace(Request("companyName")) & "'" & _
		")"
		
	' Response.Write(companySQLINSERT & "<br><br>")
	connSWPPP.Execute(companySQLINSERT)
	
	maxCompanySQLSELECT = "SELECT MAX(companyID) FROM Companies"
	
	Set rsMaxCompanyID = connSWPPP.Execute(maxCompanySQLSELECT)
	maxCompanyID = rsMaxCompanyID(0)
	
	rsMaxCompanyID.Close
	Set rsMaxCompanyID = Nothing
	
	inspectSQLINSERT = "INSERT INTO Inspections (" & _
		"inspecDate, projectName, projectAddr, projectCity, projectState, projectZip, " & _
		"projectCounty, onsiteContact, officePhone, emergencyPhone, companyID, " & _
		"reportType, inches, bmpsInPlace, sediment, userID, siteMap, compAddr, " & _
		"compAddr2, compCity, compState, compZip, compPhone, compContact, " & _
		"contactPhone, contactFax, contactEmail" & _
		") VALUES (" & _
		"'" & strQuoteReplace(Request("inspecDate")) & "'" & _
		", '" & strQuoteReplace(Request("projectName")) & "'" & _
		", '" & strQuoteReplace(Request("projectAddr")) & "'" & _
		", '" & strQuoteReplace(Request("projectCity")) & "'" & _
		", '" & Request("projectState") & "'" & _
		", '" & strQuoteReplace(Request("projectZip")) & "'" & _
		", '" & Request("projectCounty") & "'" & _
		", '" & strQuoteReplace(Request("onsiteContact")) & "'" & _
		", '" & strQuoteReplace(Request("officePhone")) & "'" & _
		", '" & strQuoteReplace(Request("emergencyPhone")) & "'" & _
		", " & maxCompanyID & _
		", '" & Request("reportType") & "'" & _
		", " & Request("inches") & _
		", " & Request("bmpsInPlace") & _
		", " & Request("sediment") & _
		", " & userID & _
		", '" & Request("siteMap") & "'" & _
		", '" & strQuoteReplace(Request("compAddr")) & "'" & _
		", '" & strQuoteReplace(Request("compAddr2")) & "'" & _
		", '" & strQuoteReplace(Request("compCity")) & "'" & _
		", '" & Request("compState") & "'" & _
		", '" & strQuoteReplace(Request("compZip")) & "'" & _
		", '" & strQuoteReplace(Request("compPhone")) & "'" & _
		", '" & strQuoteReplace(Request("compContact")) & "'" & _
		", '" & strQuoteReplace(Request("contactPhone")) & "'" & _
		", '" & strQuoteReplace(Request("contactFax")) & "'" & _
		", '" & strQuoteReplace(Request("contactEmail")) & "'" & _
		")"
		
	' Response.Write(inspectSQLINSERT & "<br><br>")
	connSWPPP.Execute(inspectSQLINSERT)
	
	coUserSQLINSERT = "INSERT INTO CompanyUsers (" & _
		"userID, companyID, rights" & _
		") VALUES (" & _
		userID & _
		", " & maxCompanyID & _
		", 'inspector'" & _
		")"
		
	' Response.Write(coUserSQLINSERT & "<br><br>")
	connSWPPP.Execute(coUserSQLINSERT)
	
	maxInspectSQLSELECT = "SELECT MAX(inspecID) FROM Inspections"
	
	Set rsMaxInspectID = connSWPPP.Execute(maxInspectSQLSELECT)
	maxInspectID = rsMaxInspectID(0)
	
	rsMaxInspectID.Close
	Set rsMaxInspectID = Nothing
	
	connSWPPP.Close
	Set connSWPPP = Nothing
	
	Response.Redirect("addCoordinate.asp?inspecID=" & maxInspectID)
	
End If
%>
<html>
<head>
<title>SWPPP INSPECTIONS : Add New Inspection Report</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
<script language="JavaScript" src="../js/validReports.js"></script>
<script language="JavaScript" src="../js/validReports1.2.js"></script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Add New Inspection Report</h1>
    
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
	<form method="post" action="<% = Request.ServerVariables("script_name") %>" onSubmit="return isReady(this)";>
		<!-- date -->
		<tr align="center"> 
			<td colspan="2"><font color="red">Duplicate Information Error!!!</font><br>
				Company &quot;<font color="red">foobar</font>&quot; already exists!<br>
				Please, use a <a href="reportSelect.asp">default report</a> or 
				change company name.</td>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr> 
			<td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Date:</b></td>
			<td bgcolor="#999999"> <input type="text" name="inspecDate" size="10" maxlength="10"> 
				<small>&nbsp;(mm / dd / yyyy)</small></td>
		</tr>
		<!-- project name -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Project Name:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectName" size="50" maxlength="50"></td>
		</tr>
		<!-- project location -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Project Location:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectAddr" size="50" maxlength="50"> 
			</td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>City, State, Zip:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectCity" size="20" maxlength="20"> 
				&nbsp; <select name="projectState">
					<option value="AK">AK</option>
					<option value="AL">AL</option>
					<option value="AR">AR</option>
					<option value="AZ">AZ</option>
					<option value="CA">CA</option>
					<option value="CO">CO</option>
					<option value="CT">CT</option>
					<option value="DC">DC</option>
					<option value="DE">DE</option>
					<option value="FL">FL</option>
					<option value="GA">GA</option>
					<option value="HI">HI</option>
					<option value="IA">IA</option>
					<option value="ID">ID</option>
					<option value="IL">IL</option>
					<option value="IN">IN</option>
					<option value="KS">KS</option>
					<option value="KY">KY</option>
					<option value="LA">LA</option>
					<option value="MA">MA</option>
					<option value="MD">MD</option>
					<option value="ME">ME</option>
					<option value="MI">MI</option>
					<option value="MN">MN</option>
					<option value="MO">MO</option>
					<option value="MS">MS</option>
					<option value="MT">MT</option>
					<option value="NC">NC</option>
					<option value="ND">ND</option>
					<option value="NE">NE</option>
					<option value="NH">NH</option>
					<option value="NJ">NJ</option>
					<option value="NM">NM</option>
					<option value="NV">NV</option>
					<option value="NY">NY</option>
					<option value="OH">OH</option>
					<option value="OK">OK</option>
					<option value="OR">OR</option>
					<option value="PA">PA</option>
					<option value="RI">RI</option>
					<option value="SC">SC</option>
					<option value="SD">SD</option>
					<option value="TN">TN</option>
					<option value="TX">TX</option>
					<option value="UT">UT</option>
					<option value="VA">VA</option>
					<option value="VT">VT</option>
					<option value="WA">WA</option>
					<option value="WI">WI</option>
					<option value="WV">WV</option>
					<option value="WY">WY</option>
				</select> &nbsp; <input type="text" name="projectZip" size="5" maxlength="5"> 
			</td>
		</tr>
		<!-- onsite contact -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>County:</b></td>
			<td bgcolor="#999999"><select name="projectCounty">
					<option value="Tarrant">Tarrant</option>
					<option value="Bosque">Bosque</option>
					<option value="Dallas">Dallas</option>
					<option value="Denton">Denton</option>
					<option value="Collin">Collin</option>
					<option value="Rockwall">Rockwall</option>
					<option value="Ellis">Ellis</option>
					<option value="Johnson">Johnson</option>
					<option value="Kaufman">Kaufman</option>
					<option value="Cooke">Cooke</option>
					<option value="Fannin">Fannin</option>
					<option value="Grayson">Grayson</option>
					<option value="Hill">Hill</option>
					<option value="Hood">Hood</option>
					<option value="Hunt">Hunt</option>
					<option value="Montague">Montague</option>
					<option value="Navarro">Navarro</option>
					<option value="Parker">Parker</option>
					<option value="Somervell">Somervell</option>
					<option value="Williamson">Williamson</option>
					<option value="Wise">Wise</option>
				</select></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="onsiteContact" size="50" maxlength="50"></td>
		</tr>
		<!-- office # -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Office Phone:</b></td>
			<td bgcolor="#999999"><input name="officePhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- emergency # -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"> <b>Emergency Phone:</b></td>
			<td bgcolor="#999999"><input name="emergencyPhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- company -->
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr>
		<tr> 
			<td colspan="2" align="center"><font size="+1">Company Information</font></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Company Name:</b></td>
			<td bgcolor="#999999"><input type="text" name="companyName" size="50" maxlength="50"></td>
		</tr>
		<!-- Address -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Address 1:</b></td>
			<td bgcolor="#999999"><input name="compAddr" type="text" size="50" maxlength="50"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Address 2:</b></td>
			<td bgcolor="#999999"><input name="compAddr2" type="text" size="50" maxlength="50"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>City:</b></td>
			<td bgcolor="#999999"><input name="compCity" type="text" size="20" maxlength="20"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>State:</b></td>
			<td bgcolor="#999999"><select name="compState">
					<option value="AK">AK</option>
					<option value="AL">AL</option>
					<option value="AR">AR</option>
					<option value="AZ">AZ</option>
					<option value="CA">CA</option>
					<option value="CO">CO</option>
					<option value="CT">CT</option>
					<option value="DC">DC</option>
					<option value="DE">DE</option>
					<option value="FL">FL</option>
					<option value="GA">GA</option>
					<option value="HI">HI</option>
					<option value="IA">IA</option>
					<option value="ID">ID</option>
					<option value="IL">IL</option>
					<option value="IN">IN</option>
					<option value="KS">KS</option>
					<option value="KY">KY</option>
					<option value="LA">LA</option>
					<option value="MA">MA</option>
					<option value="MD">MD</option>
					<option value="ME">ME</option>
					<option value="MI">MI</option>
					<option value="MN">MN</option>
					<option value="MO">MO</option>
					<option value="MS">MS</option>
					<option value="MT">MT</option>
					<option value="NC">NC</option>
					<option value="ND">ND</option>
					<option value="NE">NE</option>
					<option value="NH">NH</option>
					<option value="NJ">NJ</option>
					<option value="NM">NM</option>
					<option value="NV">NV</option>
					<option value="NY">NY</option>
					<option value="OH">OH</option>
					<option value="OK">OK</option>
					<option value="OR">OR</option>
					<option value="PA">PA</option>
					<option value="RI">RI</option>
					<option value="SC">SC</option>
					<option value="SD">SD</option>
					<option value="TN">TN</option>
					<option value="TX">TX</option>
					<option value="UT">UT</option>
					<option value="VA">VA</option>
					<option value="VT">VT</option>
					<option value="WA">WA</option>
					<option value="WI">WI</option>
					<option value="WV">WV</option>
					<option value="WY">WY</option>
				</select></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Zip:</b></td>
			<td bgcolor="#999999"><input name="compZip" type="text" size="5" maxlength="5"></td>
		</tr>
		<!-- main telephone number -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Company Phone:</b></td>
			<td bgcolor="#999999"><input name="compPhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- contact -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="compContact" size="50" maxlength="50"></td>
		</tr>
		<!-- phone -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact Phone:</b></td>
			<td bgcolor="#999999"><input name="contactPhone" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- fax -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact Fax:</b></td>
			<td bgcolor="#999999"><input name="contactFax" type="text" size="20" maxlength="20"></td>
		</tr>
		<!-- e-mail -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact E-Mail:</b></td>
			<td bgcolor="#999999"><input type="text" name="contactEmail" size="50" maxlength="50"></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2"><hr align="center" width="95%" size="1"></td>
		</tr>
		<tr> 
			<td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<!-- Type of Report? Weekly, Storm, Complaint, ? -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Type of Report:</b></td>
			<td bgcolor="#999999"><select name="reportType">
					<option value="biWeekly">Bi-Weekly</option>
					<option value="storm">Storm</option>
					<option value="complaint">Complaint</option>
					<option value="none">None</option>
				</select></td>
		</tr>
		<!-- Rain -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Inches of Rain:</b></td>
			<td bgcolor="#999999"><input type="text" name="inches" size="6" maxlength="6"></td>
		</tr>
		<!-- BMPs? y/n -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Are BMPs in place?</b></td>
			<td bgcolor="#999999"><select name="bmpsInPlace">
					<option value="1">Yes</option>
					<option value="0">No</option>
				</select></td>
		</tr>
		<!-- sediment loss or pollution? y/n -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Sediment Loss or Pollution?</b></td>
			<td bgcolor="#999999"><select name="sediment">
					<option value="1">Yes</option>
					<option value="0">No</option>
				</select></td>
		</tr>
		<%
' Admin can change inspector name.
If Session("validAdmin") Then
	
	insSQLSELECT = "SELECT DISTINCT Users.userID," & _
		" firstName, lastName" & _
		" FROM Users, CompanyUsers" & _
		" WHERE Users.userID = CompanyUsers.userID" & _
		" AND rights = 'inspector'"
	
	Set connUser = connSWPPP.execute(insSQLSELECT)
%>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>Inspector:</strong></td>
			<td bgcolor="#999999"><select name="inspector">
					<% Do While Not connUser.EOF %>
					<option value="<% = connUser("userID") %>"> 
					<% = Trim(connUser("firstName")) & "&nbsp;" & Trim(connUser("lastName")) %>
					</option>
					<%
				 	connUser.moveNext
				Loop
				
	connUser.Close
	Set connUser = Nothing
	
End If ' Session("validAdmin")
%>
				</select></td>
		</tr>
		<%
'		tempDev = "dev\"
		baseDir = "d:\vol\swpppinspections.com\www\htdocs\"
		Set folderSvrObj = Server.CreateObject("Scripting.FileSystemObject")
		
		Set objSteMapDir = folderSvrObj.GetFolder(baseDir & "images\sitemaps\")
		Set siteMapImage = objSteMapDir.Files
%>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>Site Map File:</strong></td>
			<td bgcolor="#999999" nowrap><select name="siteMap">
					<%
			For Each Item In siteMapImage
				fileName = Item.Name
%>
					<option value="<% = fileName %>"> 
					<% = fileName %>
					</option>
					<%
			Next
			
		Set objSteMapDir = Nothing
		Set siteMapImage = Nothing
		
connSWPPP.Close
Set connSWPPP = Nothing
%>
				</select> &nbsp;&nbsp; <input type="button" value="Upload Site Map File" 
					onClick="location='upSiteMapAddRprt.asp'; return false";></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="2" align="center"> <input name="submit" type="submit" value="Add New Report & Company Info"></td>
		</tr>
		<tr> 
			<td colspan="2">&nbsp;</td>
		</tr>
	</form>
</table>
</body>
</html>
