<%	Set connSWPPP = Server.CreateObject("adodb.connection")
	connSWPPP.Open "database=swppp; dsn=SWPPP; uid=SWAccess; password=C0c0nut$" 
	connSWPPP.CommandTimeout=360%>