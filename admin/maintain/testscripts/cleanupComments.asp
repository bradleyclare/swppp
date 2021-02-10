<%Response.Buffer = False

If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../../loginUser.asp")
End If 

'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout

%><!-- #INCLUDE FILE="../../connSWPPP.asp" --><%

Server.ScriptTimeout=1500

'Response.Write(Request.Form.Count & "<br>")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Cleanup Comments</TITLE>
	<LINK REL=stylesheet HREF="../../../global.css" type="text/css">
</HEAD>
<BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0 marginwidth="5" marginheight="5">
<!-- #INCLUDE FILE="../../adminHeader3.inc" -->  
<%
'get all comments
commSQLSELECT = "SELECT commentID, comment, userID, date, coID" &_
    " FROM CoordinatesComments" &_
	" WHERE projectID IS NULL" &_
    " ORDER BY date DESC"
'Response.Write(commSQLSELECT)
Set rsComm = connSWPPP.execute(commSQLSELECT)%>
    
<h1>Cleanup Comments</h1>                    
<% If rsComm.EOF Then %>
	<p>No Comments Found</p>
<% Else
	cnt = 0
    limitCnt = 50000
    updateCnt = 0
    Do While Not rsComm.EOF
        cnt = cnt + 1
		If cnt > limitCnt Then
			Exit Do
		End If
		coID = Trim(rsComm("coID"))
        commentID = Trim(rsComm("commentID"))
		
		'find item to get the inspecID 
		coordSQLSELECT = "SELECT coID, C.inspecID, projectID FROM Coordinates as C INNER JOIN Inspections I on C.inspecID = I.inspecID" &_
			" WHERE coID=" & coID
		'Response.Write(coordSQLSELECT)
		Set rsCoord = connSWPPP.execute(coordSQLSELECT)
		
		Do While Not rsCoord.EOF
			inspecID = Trim(rsCoord("inspecID"))
			projectID = Trim(rsCoord("projectID"))
			
			updateCnt = updateCnt + 1   
			inspectSQLUPDATE = "UPDATE CoordinatesComments SET" & _
			" inspecID = " & inspecID & _
			" , projectID = " & projectID & _
			" WHERE commentID = " & commentID
			response.Write(cnt & " : " & inspectSQLUPDATE & "<br/>")
			connSWPPP.Execute(inspectSQLUPDATE)
			
			rsCoord.MoveNext
		LOOP
        rsComm.MoveNext
    Loop 'RSO
    rsComm.Close
    SET rsComm=nothing
	rsCoord.Close
	SET rsCoord=nothing
End If 'end RSO.EOF %>
<h4>DONE - Total Reports: <%=cnt%>, Updates: <%=updateCnt%></h4>
</BODY>
</HTML>
<% connSWPPP.close
SET connSWPPP=nothing %>