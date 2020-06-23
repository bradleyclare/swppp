<%@ Language="VBScript" %>
<% If _
	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
    Not Session("validErosion") And _
	Not Session("validUser") _
Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")
End If

projectID = Request("pID")

%><!-- #include file="../admin/connSWPPP.asp" --><%


SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& projectID
'response.Write(SQL2)
Set RS2=connSWPPP.execute(SQL2) %>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
	<title>SWPPP INSPECTIONS : View Alert Notes</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<body>
<center>
    <img src="../images/color_logo_report.jpg" width="300"><br><br>
    <font size="+1"><b>alert items for<br/> (<%=projectID %>) <%= RS2("projectName") %>&nbsp;<%= RS2("projectPhase")%></b></font>
    <br /><br />
</center>
<center>
<table>
    <tr>
        <th width="25%">inspection date</th>
        <th width="75%">alert item</th>
    </tr>
    <% dbg_str = ""
    inspectInfoSQLSELECT = "SELECT DISTINCT inspecID, inspecDate, totalItems, completedItems, includeItems, compliance, released, systemic, systemicNote, p.projectName, p.projectPhase" & _
	" FROM Projects as p, Inspections as i" & _
	" WHERE i.projectID=p.projectID" &_
	" AND i.projectID="& projectID &_
	" ORDER BY inspecDate DESC"
    'Response.Write(inspectInfoSQLSELECT & "<br>")
    Set rsInspectInfo = connSWPPP.Execute(inspectInfoSQLSELECT) 
            
    If rsInspectInfo.EOF Then
	    Response.Write("<tr><td colspan='10' align='center'><i style='font-size: 15px'>There are no inspection reports found.</i></td></tr>")
    Else
        n = 0
        cnt = 0
        siteMapInspecID = 0
	     Do While Not rsInspectInfo.EOF   
            inspecID = rsInspectInfo("inspecID")
            inspecDate = Trim(rsInspectInfo("inspecDate"))
            includeItems = rsInspectInfo("includeItems")
            systemic = rsInspectInfo("systemic")
            systemicNote = rsInspectInfo("systemicNote")
            
            if systemic = True then %>
               <tr><td><%=inspecDate %></td>
               <td><%=systemicNote %></td></tr>
            <% end if
            
        rsInspectInfo.MoveNext
        LOOP
    End If 'end rsInspectInfo
    rsInspectInfo.Close
    SET rsInspectInfo=nothing
    connSWPPP.Close
    Set connSWPPP = Nothing %>
</table>
</center>
</body>
</html>
