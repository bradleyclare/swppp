<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

default = Request("default")

If default = "" Then Response.Redirect("reportSelect.asp") End If
%>
<!-- #include file="../connSWPPP.asp" -->
<%
companyID = Request("companyID")

If companyID <> "" Then
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function
	
	userID = Session("userID")
	inspector = strQuoteReplace(Request("inspector"))
	
	If inspector <> "" Then userID = inspector End If
	
	companySQLUPDATE = "UPDATE Companies SET" & _
		" companyName = '" & strQuoteReplace(Request("companyName")) & "'" & _
		", compAddr = '" & strQuoteReplace(Request("compAddr")) & "'" & _
		", compAddr2 = '" & strQuoteReplace(Request("compAddr2")) & "'" & _
		", compCity = '" & strQuoteReplace(Request("compCity")) & "'" & _
		", compState = '" & Request("compState") & "'" & _
		", compZip = '" & strQuoteReplace(Request("compZip")) & "'" & _
		", compPhone = '" & strQuoteReplace(Request("compPhone")) & "'" & _
		", compContact = '" & strQuoteReplace(Request("compContact")) & "'" & _
		", contactPhone = '" & strQuoteReplace(Request("contactPhone")) & "'" & _
		", contactFax = '" & strQuoteReplace(Request("contactFax")) & "'" & _
		", contactEmail = '" & strQuoteReplace(Request("contactEmail")) & "'" & _
		" WHERE companyID = " & companyID
		
	' Response.Write(companySQLUPDATE & "<br><br>")
	connSWPPP.Execute(companySQLUPDATE)
	
	inspectSQLINSERT = "INSERT INTO Inspections (" & _
		"inspecDate, projectName, projectAddr, projectCity, projectState, projectZip, " & _
		"projectCounty, onsiteContact, officePhone, emergencyPhone, companyID, " & _
		"reportType, inches, bmpsInPlace, sediment, userID, siteMap" & _
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
		", " & companyID & _
		", '" & Request("reportType") & "'" & _
		", " & Request("inches") & _
		", " & Request("bmpsInPlace") & _
		", " & Request("sediment") & _
		", " & userID & _
		", '" & Request("siteMap") & "'" & _
		")"
		
	' Response.Write(inspectSQLINSERT & "<br><br>")
	connSWPPP.Execute(inspectSQLINSERT)
	
	maxInspectSQLSELECT = "SELECT MAX(inspecID) FROM Inspections"
	
	Set rsMaxInspectID = connSWPPP.Execute(maxInspectSQLSELECT)
	maxInspectID = rsMaxInspectID(0)
	
	rsMaxInspectID.Close
	Set rsMaxInspectID = Nothing
	
	coordCount = Request("coordCount")
	
	For i = 1 To coordCount
		coordinates = strQuoteReplace(Request("coordinates" & i))
		existingBMP = strQuoteReplace(Request("existingBMP" & i))
		correctiveMods = strQuoteReplace(Request("correctiveMods" & i))
		
		coordSQLINSERT = "INSERT INTO Coordinates (" & _
			"inspecID, coordinates, existingBMP, correctiveMods" & _
			") VALUES (" & _
			maxInspectID & _
			", '" & coordinates & "'" & _
			", '" & existingBMP & "'" & _
			", '" & correctiveMods & "'" & _
			")"
			
		' Response.Write(coordSQLINSERT & "<br><br>")
		connSWPPP.Execute(coordSQLINSERT)
		
	Next
	
	connSWPPP.Close
	Set connSWPPP = Nothing
	
	Response.Redirect("viewReports.asp")
	
End If

idArray = Split(default, "~")
inspecID = idArray(0)
companyID = idArray(1)

reportSQLSELECT = "SELECT inspecDate, projectName, projectAddr, projectCity, projectState" & _
	", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, companyName" & _
	", compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact" & _
	", contactPhone, contactFax, contactEmail, reportType, inches, bmpsInPlace" & _
	", sediment, userID, siteMap" & _
	" FROM Inspections, Companies" & _
	" WHERE inspecID = " & inspecID & _
	" AND Inspections.companyID = " & companyID
	
' Response.Write(reportSQLSELECT & "<br>")
Set connReport = connSWPPP.execute(reportSQLSELECT)
%>
<html>
<head>
<title>SWPPP INSPECTIONS : Add Inspection Report</title>
<link rel="stylesheet" type="text/css" href="../../global.css">
<script language="JavaScript" src="../js/validReports.js"></script>
<script language="JavaScript" src="../js/validReports1.2.js"></script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Add Inspection Report</h1>
<form method="post" action="<% = Request.ServerVariables("script_name") %>" onSubmit="return isReady(this)";>
    <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0">
		<input type="hidden" name="default" value="True">
		<input type="hidden" name="companyID" value="<% = companyID %>">
		<!-- date -->
		<tr> 
			<td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Date:</b></td>
			<td bgcolor="#999999"> <input type="text" name="inspecDate" size="10" maxlength="10"
					value="<% = Trim(connReport("inspecDate")) %>"> <small>&nbsp;(mm 
				/ dd / yyyy)</small></td>
		</tr>
		<!-- project name -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Project Name:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectName" size="50" maxlength="50" 
				value="<% = Trim(connReport("projectName")) %>"></td>
		</tr>
		<!-- project location -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Project Location:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectAddr" size="50" maxlength="50" 
				value="<% = Trim(connReport("projectAddr")) %>"> </td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>City, State, Zip:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectCity" size="20" maxlength="20" 
				value="<% = Trim(connReport("projectCity")) %>"> &nbsp; <select name="projectState">
					<option value="AK"<% If Trim(connReport("projectState")) = "AK" Then %> selected<% End If %>>AK</option>
					<option value="AL"<% If Trim(connReport("projectState")) = "AL" Then %> selected<% End If %>>AL</option>
					<option value="AR"<% If Trim(connReport("projectState")) = "AR" Then %> selected<% End If %>>AR</option>
					<option value="AZ"<% If Trim(connReport("projectState")) = "AZ" Then %> selected<% End If %>>AZ</option>
					<option value="CA"<% If Trim(connReport("projectState")) = "CA" Then %> selected<% End If %>>CA</option>
					<option value="CO"<% If Trim(connReport("projectState")) = "CO" Then %> selected<% End If %>>CO</option>
					<option value="CT"<% If Trim(connReport("projectState")) = "CT" Then %> selected<% End If %>>CT</option>
					<option value="DC"<% If Trim(connReport("projectState")) = "DC" Then %> selected<% End If %>>DC</option>
					<option value="DE"<% If Trim(connReport("projectState")) = "DE" Then %> selected<% End If %>>DE</option>
					<option value="FL"<% If Trim(connReport("projectState")) = "FL" Then %> selected<% End If %>>FL</option>
					<option value="GA"<% If Trim(connReport("projectState")) = "GA" Then %> selected<% End If %>>GA</option>
					<option value="HI"<% If Trim(connReport("projectState")) = "HI" Then %> selected<% End If %>>HI</option>
					<option value="IA"<% If Trim(connReport("projectState")) = "IA" Then %> selected<% End If %>>IA</option>
					<option value="ID"<% If Trim(connReport("projectState")) = "ID" Then %> selected<% End If %>>ID</option>
					<option value="IL"<% If Trim(connReport("projectState")) = "IL" Then %> selected<% End If %>>IL</option>
					<option value="IN"<% If Trim(connReport("projectState")) = "IN" Then %> selected<% End If %>>IN</option>
					<option value="KS"<% If Trim(connReport("projectState")) = "KS" Then %> selected<% End If %>>KS</option>
					<option value="KY"<% If Trim(connReport("projectState")) = "KY" Then %> selected<% End If %>>KY</option>
					<option value="LA"<% If Trim(connReport("projectState")) = "LA" Then %> selected<% End If %>>LA</option>
					<option value="MA"<% If Trim(connReport("projectState")) = "MA" Then %> selected<% End If %>>MA</option>
					<option value="MD"<% If Trim(connReport("projectState")) = "MD" Then %> selected<% End If %>>MD</option>
					<option value="ME"<% If Trim(connReport("projectState")) = "ME" Then %> selected<% End If %>>ME</option>
					<option value="MI"<% If Trim(connReport("projectState")) = "MI" Then %> selected<% End If %>>MI</option>
					<option value="MN"<% If Trim(connReport("projectState")) = "MN" Then %> selected<% End If %>>MN</option>
					<option value="MO"<% If Trim(connReport("projectState")) = "MO" Then %> selected<% End If %>>MO</option>
					<option value="MS"<% If Trim(connReport("projectState")) = "MS" Then %> selected<% End If %>>MS</option>
					<option value="MT"<% If Trim(connReport("projectState")) = "MT" Then %> selected<% End If %>>MT</option>
					<option value="NC"<% If Trim(connReport("projectState")) = "NC" Then %> selected<% End If %>>NC</option>
					<option value="ND"<% If Trim(connReport("projectState")) = "ND" Then %> selected<% End If %>>ND</option>
					<option value="NE"<% If Trim(connReport("projectState")) = "NE" Then %> selected<% End If %>>NE</option>
					<option value="NH"<% If Trim(connReport("projectState")) = "NH" Then %> selected<% End If %>>NH</option>
					<option value="NJ"<% If Trim(connReport("projectState")) = "NJ" Then %> selected<% End If %>>NJ</option>
					<option value="NM"<% If Trim(connReport("projectState")) = "NM" Then %> selected<% End If %>>NM</option>
					<option value="NV"<% If Trim(connReport("projectState")) = "NV" Then %> selected<% End If %>>NV</option>
					<option value="NY"<% If Trim(connReport("projectState")) = "NY" Then %> selected<% End If %>>NY</option>
					<option value="OH"<% If Trim(connReport("projectState")) = "OH" Then %> selected<% End If %>>OH</option>
					<option value="OK"<% If Trim(connReport("projectState")) = "OK" Then %> selected<% End If %>>OK</option>
					<option value="OR"<% If Trim(connReport("projectState")) = "OR" Then %> selected<% End If %>>OR</option>
					<option value="PA"<% If Trim(connReport("projectState")) = "PA" Then %> selected<% End If %>>PA</option>
					<option value="RI"<% If Trim(connReport("projectState")) = "RI" Then %> selected<% End If %>>RI</option>
					<option value="SC"<% If Trim(connReport("projectState")) = "SC" Then %> selected<% End If %>>SC</option>
					<option value="SD"<% If Trim(connReport("projectState")) = "SD" Then %> selected<% End If %>>SD</option>
					<option value="TN"<% If Trim(connReport("projectState")) = "TN" Then %> selected<% End If %>>TN</option>
					<option value="TX"<% If Trim(connReport("projectState")) = "TX" OR Trim(connReport("projectState")) = "" Then %> selected<% End If %>>TX</option>
					<option value="UT"<% If Trim(connReport("projectState")) = "UT" Then %> selected<% End If %>>UT</option>
					<option value="VA"<% If Trim(connReport("projectState")) = "VA" Then %> selected<% End If %>>VA</option>
					<option value="VT"<% If Trim(connReport("projectState")) = "VT" Then %> selected<% End If %>>VT</option>
					<option value="WA"<% If Trim(connReport("projectState")) = "WA" Then %> selected<% End If %>>WA</option>
					<option value="WI"<% If Trim(connReport("projectState")) = "WI" Then %> selected<% End If %>>WI</option>
					<option value="WV"<% If Trim(connReport("projectState")) = "WV" Then %> selected<% End If %>>WV</option>
					<option value="WY"<% If Trim(connReport("projectState")) = "WY" Then %> selected<% End If %>>WY</option>
				</select> &nbsp; <input type="text" name="projectZip" size="5" maxlength="5" 
					value="<% = Trim(connReport("projectZip")) %>"> </td>
		</tr>
		<!-- onsite contact -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>County:</b></td>
			<td bgcolor="#999999"><select name="projectCounty">
					<option value="Tarrant"<% If Trim(connReport("projectState")) = "Tarrant" Then %> selected<% End If %>>Tarrant</option>
					<option value="Bosque"<% If Trim(connReport("projectCounty")) = "Bosque" Then %> selected<% End If %>>Bosque</option>
					<option value="Dallas"<% If Trim(connReport("projectCounty")) = "Dallas" Then %> selected<% End If %>>Dallas</option>
					<option value="Denton"<% If Trim(connReport("projectCounty")) = "Denton" Then %> selected<% End If %>>Denton</option>
					<option value="Collin"<% If Trim(connReport("projectCounty")) = "Collin" Then %> selected<% End If %>>Collin</option>
					<option value="Rockwall"<% If Trim(connReport("projectCounty")) = "Rockwall" Then %> selected<% End If %>>Rockwall</option>
					<option value="Ellis"<% If Trim(connReport("projectCounty")) = "Ellis" Then %> selected<% End If %>>Ellis</option>
					<option value="Johnson"<% If Trim(connReport("projectCounty")) = "Johnson" Then %> selected<% End If %>>Johnson</option>
					<option value="Kaufman"<% If Trim(connReport("projectCounty")) = "Kaufman" Then %> selected<% End If %>>Kaufman</option>
					<option value="Cooke"<% If Trim(connReport("projectCounty")) = "Cooke" Then %> selected<% End If %>>Cooke</option>
					<option value="Fannin"<% If Trim(connReport("projectCounty")) = "Fannin" Then %> selected<% End If %>>Fannin</option>
					<option value="Grayson"<% If Trim(connReport("projectCounty")) = "Grayson" Then %> selected<% End If %>>Grayson</option>
					<option value="Hill"<% If Trim(connReport("projectCounty")) = "Hill" Then %> selected<% End If %>>Hill</option>
					<option value="Hood"<% If Trim(connReport("projectCounty")) = "Hood" Then %> selected<% End If %>>Hood</option>
					<option value="Hunt"<% If Trim(connReport("projectCounty")) = "Hunt" Then %> selected<% End If %>>Hunt</option>
					<option value="Montague"<% If Trim(connReport("projectCounty")) = "Montague" Then %> selected<% End If %>>Montague</option>
					<option value="Navarro"<% If Trim(connReport("projectCounty")) = "Navarro" Then %> selected<% End If %>>Navarro</option>
					<option value="Parker"<% If Trim(connReport("projectCounty")) = "Parker" Then %> selected<% End If %>>Parker</option>
					<option value="Somervell"<% If Trim(connReport("projectCounty")) = "Somervell" Then %> selected<% End If %>>Somervell</option>
					<option value="Williamson"<% If Trim(connReport("projectCounty")) = "Williamson" Then %> selected<% End If %>>Williamson</option>
					<option value="Wise"<% If Trim(connReport("projectCounty")) = "Wise" Then %> selected<% End If %>>Wise</option>
				</select></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="onsiteContact" size="50" maxlength="50" 
				value="<% = Trim(connReport("onsiteContact")) %>"></td>
		</tr>
		<!-- office # -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Office Phone:</b></td>
			<td bgcolor="#999999"><input name="officePhone" type="text" size="20" maxlength="20"
				 value="<% = Trim(connReport("officePhone")) %>"></td>
		</tr>
		<!-- emergency # -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"> <b>Emergency Phone:</b></td>
			<td bgcolor="#999999"><input name="emergencyPhone" type="text" size="20" maxlength="20"
				 value="<% = Trim(connReport("emergencyPhone")) %>"></td>
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
		<tr align="center"> 
			<td colspan="2"><font size="+1">Company Information</font></td>
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
			<td bgcolor="#999999"><input type="text" name="companyName" size="50" maxlength="50"
				 value="<% = Trim(connReport("companyName")) %>"></td>
		</tr>
		<!-- Address -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Address 1:</b></td>
			<td bgcolor="#999999"><input name="compAddr" type="text" size="50" maxlength="50" 
				value="<% = Trim(connReport("compAddr")) %>"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Address 2:</b></td>
			<td bgcolor="#999999"><input name="compAddr2" type="text" size="50" maxlength="50" 
				value="<% = Trim(connReport("compAddr2")) %>"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>City:</b></td>
			<td bgcolor="#999999"><input name="compCity" type="text" size="20" maxlength="20" 
				value="<% = Trim(connReport("compCity")) %>"></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>State:</b></td>
			<td bgcolor="#999999"><select name="compState">
					<option value="AK"<% If Trim(connReport("compState")) = "AK" Then %> selected<% End If %>>AK</option>
					<option value="AL"<% If Trim(connReport("compState")) = "AL" Then %> selected<% End If %>>AL</option>
					<option value="AR"<% If Trim(connReport("compState")) = "AR" Then %> selected<% End If %>>AR</option>
					<option value="AZ"<% If Trim(connReport("compState")) = "AZ" Then %> selected<% End If %>>AZ</option>
					<option value="CA"<% If Trim(connReport("compState")) = "CA" Then %> selected<% End If %>>CA</option>
					<option value="CO"<% If Trim(connReport("compState")) = "CO" Then %> selected<% End If %>>CO</option>
					<option value="CT"<% If Trim(connReport("compState")) = "CT" Then %> selected<% End If %>>CT</option>
					<option value="DC"<% If Trim(connReport("compState")) = "DC" Then %> selected<% End If %>>DC</option>
					<option value="DE"<% If Trim(connReport("compState")) = "DE" Then %> selected<% End If %>>DE</option>
					<option value="FL"<% If Trim(connReport("compState")) = "FL" Then %> selected<% End If %>>FL</option>
					<option value="GA"<% If Trim(connReport("compState")) = "GA" Then %> selected<% End If %>>GA</option>
					<option value="HI"<% If Trim(connReport("compState")) = "HI" Then %> selected<% End If %>>HI</option>
					<option value="IA"<% If Trim(connReport("compState")) = "IA" Then %> selected<% End If %>>IA</option>
					<option value="ID"<% If Trim(connReport("compState")) = "ID" Then %> selected<% End If %>>ID</option>
					<option value="IL"<% If Trim(connReport("compState")) = "IL" Then %> selected<% End If %>>IL</option>
					<option value="IN"<% If Trim(connReport("compState")) = "IN" Then %> selected<% End If %>>IN</option>
					<option value="KS"<% If Trim(connReport("compState")) = "KS" Then %> selected<% End If %>>KS</option>
					<option value="KY"<% If Trim(connReport("compState")) = "KY" Then %> selected<% End If %>>KY</option>
					<option value="LA"<% If Trim(connReport("compState")) = "LA" Then %> selected<% End If %>>LA</option>
					<option value="MA"<% If Trim(connReport("compState")) = "MA" Then %> selected<% End If %>>MA</option>
					<option value="MD"<% If Trim(connReport("compState")) = "MD" Then %> selected<% End If %>>MD</option>
					<option value="ME"<% If Trim(connReport("compState")) = "ME" Then %> selected<% End If %>>ME</option>
					<option value="MI"<% If Trim(connReport("compState")) = "MI" Then %> selected<% End If %>>MI</option>
					<option value="MN"<% If Trim(connReport("compState")) = "MN" Then %> selected<% End If %>>MN</option>
					<option value="MO"<% If Trim(connReport("compState")) = "MO" Then %> selected<% End If %>>MO</option>
					<option value="MS"<% If Trim(connReport("compState")) = "MS" Then %> selected<% End If %>>MS</option>
					<option value="MT"<% If Trim(connReport("compState")) = "MT" Then %> selected<% End If %>>MT</option>
					<option value="NC"<% If Trim(connReport("compState")) = "NC" Then %> selected<% End If %>>NC</option>
					<option value="ND"<% If Trim(connReport("compState")) = "ND" Then %> selected<% End If %>>ND</option>
					<option value="NE"<% If Trim(connReport("compState")) = "NE" Then %> selected<% End If %>>NE</option>
					<option value="NH"<% If Trim(connReport("compState")) = "NH" Then %> selected<% End If %>>NH</option>
					<option value="NJ"<% If Trim(connReport("compState")) = "NJ" Then %> selected<% End If %>>NJ</option>
					<option value="NM"<% If Trim(connReport("compState")) = "NM" Then %> selected<% End If %>>NM</option>
					<option value="NV"<% If Trim(connReport("compState")) = "NV" Then %> selected<% End If %>>NV</option>
					<option value="NY"<% If Trim(connReport("compState")) = "NY" Then %> selected<% End If %>>NY</option>
					<option value="OH"<% If Trim(connReport("compState")) = "OH" Then %> selected<% End If %>>OH</option>
					<option value="OK"<% If Trim(connReport("compState")) = "OK" Then %> selected<% End If %>>OK</option>
					<option value="OR"<% If Trim(connReport("compState")) = "OR" Then %> selected<% End If %>>OR</option>
					<option value="PA"<% If Trim(connReport("compState")) = "PA" Then %> selected<% End If %>>PA</option>
					<option value="RI"<% If Trim(connReport("compState")) = "RI" Then %> selected<% End If %>>RI</option>
					<option value="SC"<% If Trim(connReport("compState")) = "SC" Then %> selected<% End If %>>SC</option>
					<option value="SD"<% If Trim(connReport("compState")) = "SD" Then %> selected<% End If %>>SD</option>
					<option value="TN"<% If Trim(connReport("compState")) = "TN" Then %> selected<% End If %>>TN</option>
					<option value="TX"<% If Trim(connReport("compState")) = "TX" Or Trim(connReport("compState")) = "" Then %> selected<% End If %>>TX</option>
					<option value="UT"<% If Trim(connReport("compState")) = "UT" Then %> selected<% End If %>>UT</option>
					<option value="VA"<% If Trim(connReport("compState")) = "VA" Then %> selected<% End If %>>VA</option>
					<option value="VT"<% If Trim(connReport("compState")) = "VT" Then %> selected<% End If %>>VT</option>
					<option value="WA"<% If Trim(connReport("compState")) = "WA" Then %> selected<% End If %>>WA</option>
					<option value="WI"<% If Trim(connReport("compState")) = "WI" Then %> selected<% End If %>>WI</option>
					<option value="WV"<% If Trim(connReport("compState")) = "WV" Then %> selected<% End If %>>WV</option>
					<option value="WY"<% If Trim(connReport("compState")) = "WY" Then %> selected<% End If %>>WY</option>
				</select></td>
		</tr>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Zip:</b></td>
			<td bgcolor="#999999"><input name="compZip" type="text" size="5" maxlength="5" 
				value="<% = Trim(connReport("compZip")) %>"></td>
		</tr>
		<!-- main telephone number -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Company Phone:</b></td>
			<td bgcolor="#999999"><input name="compPhone" type="text" size="20" maxlength="20"
				 value="<% = Trim(connReport("compPhone")) %>"></td>
		</tr>
		<!-- contact -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="compContact" size="50" maxlength="50"
				value="<% = Trim(connReport("compContact")) %>"></td>
		</tr>
		<!-- phone -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact Phone:</b></td>
			<td bgcolor="#999999"><input name="contactPhone" type="text" size="20" maxlength="20"
				value="<% = Trim(connReport("contactPhone")) %>"></td>
		</tr>
		<!-- fax -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact Fax:</b></td>
			<td bgcolor="#999999"><input name="contactFax" type="text" size="20" maxlength="20"
				value="<% = Trim(connReport("contactFax")) %>"></td>
		</tr>
		<!-- e-mail -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Contact E-Mail:</b></td>
			<td bgcolor="#999999"><input type="text" name="contactEmail" size="50" maxlength="50"
				value="<% = Trim(connReport("contactEmail")) %>"></td>
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
					<option value="biWeekly"<% If Trim(connReport("reportType")) = "biWeekly" Then %> selected<% End If %>>Bi-Weekly</option>
					<option value="storm"<% If Trim(connReport("reportType")) = "storm" Then %> selected<% End If %>>Storm</option>
					<option value="complaint"<% If Trim(connReport("reportType")) = "complaint" Then %> selected<% End If %>>Complaint</option>
					<option value="none"<% If Trim(connReport("reportType")) = "none" Then %> selected<% End If %>>None</option>
				</select></td>
		</tr>
		<!-- Rain -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Inches of Rain:</b></td>
			<td bgcolor="#999999"><input type="text" name="inches" size="6" maxlength="6"
				value="<% = Trim(connReport("inches")) %>"></td>
		</tr>
		<!-- BMPs? y/n -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Are BMPs in place?</b></td>
			<td bgcolor="#999999"><select name="bmpsInPlace">
					<option value="1"<% If connReport("bmpsInPlace") Then %> selected<% End If %>>Yes</option>
					<option value="0"<% If Not connReport("bmpsInPlace") Then %> selected<% End If %>>No</option>
				</select></td>
		</tr>
		<!-- sediment loss or pollution? y/n -->
		<tr> 
			<td align="right" bgcolor="#eeeeee"><b>Sediment Loss or Pollution?</b></td>
			<td bgcolor="#999999"><select name="sediment">
					<option value="1"<% If connReport("sediment") Then %> selected<% End If %>>Yes</option>
					<option value="0"<% If Not connReport("sediment") Then %> selected<% End If %>>No</option>
				</select></td>
		</tr>
<%
'admin can change inspector name
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
				<% Do While Not connUser.eof %>
				<option value="<% = connUser("userID") %>"
					<% If connReport("userID") = connUser("userID") Then %> selected<% End If %>> 
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
			tempDev = "dev\"
			baseDir = "C:\Inetpub\wwwroot\pwp\swppp\"
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
					<option value="<% = fileName %>"<% If Trim(connReport("siteMap")) = fileName Then %> selected<% End If %>> 
					<% = fileName %>
					</option>
<%
				Next
				
			Set objSteMapDir = Nothing
			Set siteMapImage = Nothing
			
			connReport.Close
			Set connReport = Nothing
%>
				</select> &nbsp;&nbsp; <input type="button" value="Upload Site Map File" 
					onClick="location='upSiteMapAddRprt.asp'; return false";></td>
		</tr>
		<tr> 
			<td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
	</table>
<h1>Location</h1>
    <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
		<tr> 
			<td colspan="3"><hr align="center" width="100%" size="1"></td>
		</tr>
		<%
coordCount = 0

coordSQLSELECT = "SELECT coordinates, existingBMP, correctiveMods, orderby" & _
	" FROM Coordinates" & _
	" WHERE inspecID = " & inspecID & _
	" ORDER BY orderby"
	
Set rsCoord = connSWPPP.Execute(coordSQLSELECT)

If rsCoord.EOF Then
	Response.Write("<tr><td colspan='2' align='center'><i><b>There is " & _
		"no coordinate data associated at this time.</b></i></td></tr>" & _
		"<td colspan='2'><hr align='center' width='100%' size='1'>")
		
	submitBtn = "Add Report & Associate Company Info"
	
Else
	submitBtn = "Add Report, Associate Company & Locations"
	
	Do While Not rsCoord.EOF
		correctiveMods = Trim(rsCoord("correctiveMods"))
		orderby = rsCoord("orderby")
		coordinates = Trim(rsCoord("coordinates"))
		existingBMP = Trim(rsCoord("existingBMP"))
		
		coordCount = coordCount + 1
%>
		<tr> 
			<td width="10%" rowspan="3" align="center">
<% = orderby %>
			</td>
			<td width="20%" align="right"><b>Locations:</b></td>
			<td width="70%">
				<% = coordinates %>
			</td>
			<input type="hidden" name="coordinates<% = coordCount %>" value="<% = coordinates %>">
		</tr>
		<tr> 
			<td align="right"><b>Existing BMP:</b></td>
			<td><% = existingBMP %></td>
			<input type="hidden" name="existingBMP<% = coordCount %>" value="<% = existingBMP %>">
		</tr>
		<tr> 
			<td align="right" valign="top" nowrap><b>Corrective Mods:</b></td>
			<td><% = correctiveMods %></td>
			<input type="hidden" name="correctiveMods<% = coordCount %>" value="<% = correctiveMods %>">
		</tr>
		<tr> 
			<td colspan="3"><hr align="center" width="100%" size="1"></td>
		</tr>
		<%
		rsCoord.MoveNext
	Loop
	
End If

rsCoord.Close 
Set rsCoord = Nothing

connSWPPP.Close 
Set connSWPPP = Nothing
%>
		<tr> 
			<td colspan="3">&nbsp;</td>
		</tr>
		<tr> 
			<td colspan="3" align="center"> <input name="submit" type="submit" value="<% = submitBtn %>"> 
			</td>
		</tr>
		<tr> 
			<td colspan="3">&nbsp;</td>
		</tr>
	</table>
	<input type="hidden" name="coordCount" value="<% = coordCount %>">
</form>
</body>
</html>
