<%@ Language="VBScript" %><%

If _
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

coID = Request("coID")

%><!-- #include file="../admin/connSWPPP.asp" --><%

commSQLSELECT = "SELECT comment, userID, date" &_
    " FROM CoordinatesComments WHERE coID=" & coID	
Set rsComm = connSWPPP.execute(commSQLSELECT) 

coordSQLSELECT = "SELECT coordinates, correctiveMods, assignDate, completeDate, status, repeat, useAddress, address, locationName, infoOnly, LD" &_
    " FROM Coordinates WHERE coID=" & coID
Set rsCoord = connSWPPP.execute(coordSQLSELECT) 
    
currentDate = date()

correctiveMods = Trim(rsCoord("correctiveMods"))
coordinates = Trim(rsCoord("coordinates"))
assignDate = rsCoord("assignDate")
if assignDate = "" Then
	age = "?"
Else
	age = datediff("d",assignDate,currentDate) 
End If
useAddress = rsCoord("useAddress")
address = TRIM(rsCoord("address"))
locationName = TRIM(rsCoord("locationName"))
LD = rsCoord("LD")
If LD = True Then
    correctiveMods = "(LD) " & correctiveMods
End If %>

<html>
<head>
<title>SWPPP INSPECTIONS - Open Item Comments</title>
<link rel="stylesheet" type="text/css" href="../global.css">
<link href="../css/jquery-ui.min.css" rel="stylesheet" type="text/css"/>
<link href="../css/jquery-ui.structure.min.css" rel="stylesheet" type="text/css"/>
<link href="../css/jquery-ui.theme.min.css" rel="stylesheet" type="text/css"/>
<script src="../js/jquery.js" type="text/javascript"></script>
<script src="../js/jquery-ui.min.js" type="text/javascript"></script>
<style>
    .head{ color: #808080; }
</style>
</head>

<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
    <center>
    <img src="../images/color_logo_report.jpg" width="300"><br><br>
    <font size="+1"><b>Comments for Open Item <%=coID %></b></font>
    <h3>AGE: <span class="head"><%= age %> days</span>, ASSIGN DATE: <span class="head"><%=assignDate %></span></h3> 
	<h3>LOCATION: <span class="head">
    <% if (useAddress) = False Then %>
		<%=coordinates%>
	<% Else %>
		<%=locationName%> (<%=address%>)
	<% End If %>
	</span></h3>
	<h3>ACTION ITEM: <span class="head"><%= correctiveMods %></span></h3>

    </center>
    <table cellpadding="2" cellspacing="0" border="0" width="100%">
	    <tr><th width="10%" align="left">Date</th>
            <th width="10%" align="left">User</th>
            <th width="80%" align="left">Comment</th>
	    </tr>
    <% If rsComm.EOF Then
	    Response.Write("<tr><td colspan='3' align='center'><i style='font-size: 15px'>There are no comments found.</i></td></tr>")
    Else
        n = 0
	    Do While Not rsComm.EOF   
            userID      = rsComm("userID")
            comment     = Trim(rsComm("comment"))
            commentDate = rsComm("date") 
            
            SQLSELECT = "SELECT firstName, lastName FROM Users WHERE userID = " & userID
            'Response.Write(SQLSELECT & "<br>")
            Set connUsers = connSWPPP.Execute(SQLSELECT)
            firstName = connUsers("firstName")
            lastName  = connUsers("lastName") %>
            
            <tr><td><%=commentDate %></td><td><%=firstName %><nbsp /><%=lastName %></td><td><%=comment %></td></tr>
            
            <% rsComm.MoveNext
         LOOP 'loop inpection reports
    End If %>
    </table>
    <br /><br />
</body>
</html>

<% connSWPPP.Close
SET connSWPPP=nothing %>
	