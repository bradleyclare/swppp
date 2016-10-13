<%	
	Set connSWPPP = Server.CreateObject("adodb.connection")

	On Error Resume Next
	connSWPPP.Open "database=swp3org; dsn=SWPPP; uid=swaccess; password=4sr%^Tg7h" 

	If Err.Number <> 0 Then
		Response.Write Err.Number & "<br>" & Err.Description
		Response.End
	End If

	Response.Write "Connection Opened"

	connSWPPP.Close

%>