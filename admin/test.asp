<%	
	Set connSWPPP = Server.CreateObject("adodb.connection")

	On Error Resume Next
	connSWPPP.Open "database=swppp; dsn=SWPPP; uid=SWAccess; password=iuser" 

	If Err.Number <> 0 Then
		Response.Write Err.Number & "<br>" & Err.Description
		Response.End
	End If

	Response.Write "Connection Opened"

	connSWPPP.Close

%>