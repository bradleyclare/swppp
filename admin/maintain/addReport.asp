<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

If Request("default")="" Then Response.Redirect("reportSelect.asp") End If
%> <!-- #include file="../connSWPPP.asp" --> <%

SQL0="IF OBJECT_ID('tempdb..#tmp') IS NOT NULL "&_
    " Begin DROP TABLE #tmp End" &_
    " Create Table #tmp(inspectID int, projectID int, newInspecID int); " 
listArray = Split(Request("default"),",")
for n = 0 to UBound(listArray) step 1
    idArray = Split(Trim(listArray(n)), "~")            
    inspecID = Trim(idArray(0))
    projectID = Trim(idArray(1))
    SQL0 = SQL0 +" INSERT INTO #tmp Select "& inspecID &", "& projectID &", null; " 
next 

%><!-- #include file="../connSWPPP.asp" --><%
SQL0= SQL0 &" INSERT INTO Inspections (inspecDate, projectname, projectphase, projectaddr, projectcity, projectstate, projectzip, projectcounty, onsitecontact,  " &_
	" officephone, emergencyphone, companyid, reporttype, inches, bmpsinplace, sediment, userid, compaddr, compaddr2, compcity, compstate, compzip, compphone,  " &_
	" compcontact, contactphone, contactfax, contactemail, projectid, compname, narrative, released, includeItems, compliance, totalItems, completedItems, sentRepeatItemReport, openItemAlert, groupName)  " &_
    " SELECT inspecDate='"& Date() &"', p.projectName, p.projectPhase, projectAddr, projectCity, projectState,  " &_
	" projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, companyID,  " &_
	" reportType = case when i.reportType = 'Initial' Then 'Weekly' Else i.reportType end, inches=-1, bmpsInPlace=-1, sediment=-1, userID,  " &_
	" compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax, contactEmail, p.projectID, compName, narrative, released=0, " &_
	" includeItems=1, compliance, totalItems, completedItems=0, sentRepeatItemReport=0, openItemAlert, groupName" &_
    " FROM Inspections i  " &_
	" inner join #tmp t on i.inspecID = t.inspectID and i.projectid = t.projectid" &_
	" inner join Projects p on t.projectid = p.projectid;  " &_
	" Update #tmp set newInspecID = i.InspecID " &_
    " From Inspections i inner join #tmp t on i.projectID = t.projectID " &_
    " Where i.inspecID = (select MAX(inspecID) From Inspections Where projectID = t.projectID) " &_
    " INSERT INTO OptionalImages SELECT oi.oImageName, oi.oImageDesc, oi.oImageFileName, oi.oitID, inspecID= t.newInspecID" &_
	" , oi.oOrder FROM OptionalImages oi inner join #tmp t on oi.inspecID = t.inspectID ;" &_
	" INSERT INTO Coordinates SELECT inspecID= t.newInspecID, c.coordinates, c.existingBMP, c.correctiveMods, c.orderby, c.assignDate, c.completeDate, status=0, repeat=0, c.useAddress, c.address, c.locationName, c.infoOnly, c.LD, c.NLN" &_
	" FROM Coordinates c inner join #tmp t on c.inspecID = t.inspectID;"
Response.Write(SQL0)
'response.End
connSWPPP.execute(SQL0)

'get new inspecID
'SQL = "SELECT TOP 1 * FROM Inspections"
'Set rsInspec = connSWPPP.execute(SQL)
'newInspecID = rsInspec("inspecID")

'reset completed and repeat states
'coordSQLSELECT = "SELECT coID FROM Coordinates WHERE inspecID=" & inspecID 
'Response.Write(coordSQLSELECT)
'Set rsCoord = connSWPPP.execute(coordSQLSELECT)

'Do While Not rsCoord.EOF
'	Response.Write("<br/>" & rsCoord("coID"))
'	coordSQLUPDATE = "UPDATE Coordinates SET status=0, repeat=0 WHERE coID=" & rsCoord("coID")
'	rsCoord.MoveNext
'Loop

Response.redirect("viewReports.asp") %>