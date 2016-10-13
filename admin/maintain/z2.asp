<%	inspecID=Request("ID")
	tempTrimmed=Request("temp")
msgBox(inspecID &":"& tempTrimmed)
	SQLa="sp_oImagesByType "& inspecID &",'"& tempTrimmed &"'" 
Response.Write(SQLa &"<br>")
	SET RSa=connSWPPP.execute(SQLa)	%>
<% 	DO WHILE NOT(RSa.EOF) %>
				<OPTION value="<%= RSa("oImageFileName")%>"><%= RSa("oImageFileName")%></OPTION>
<%		RSa.MoveNext
	LOOP%>
				</SELECT><BUTTON>delete</BUTTON>
					<BUTTON>add</BUTTON><select name="<%=tempTrimmed%>UP">
<%	For Each Item In TempImage
		shortName = Item.Name %>
				<option value="<% = shortName %>"><%= shortName %></option>
<%	Next 
	Set objTemp = Nothing
	Set TempImage = Nothing %>