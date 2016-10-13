<%'--Get narrative text from database and insert into page

If NOT(Session("validAdmin") OR Session("validInspector")) Then
	window.close()
end if
	IF NOT(IsNULL(Request("inspecID"))) THEN Session("inspecID")=Request("inspecID") END IF
	inspecID= Session("inspecID")
%> <!-- #include virtual="admin/connSWPPP.asp" --> <%
	If Request.Form.Count > 0 Then
		SQL1="UPDATE Inspections SET narrative='"& REPLACE(Request("narrative"),"'","#@#") &"'" &_
			" WHERE inspecID='"& inspecID &"'"
		connSWPPP.execute(SQL1)
	end if
	SQL0="SELECT narrative FROM Inspections WHERE inspecID='"& inspecID &"'"
	SET RS0= connSWPPP.execute(SQL0)
	IF RS0.EOF THEN narrative="" ELSE narrative= TRIM(RS0("narrative")) END IF
	IF IsNull(narrative) THEN narrative="" END IF %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>SWPPP INSPECTIONS : Admin : Edit Narrative</title>
<LINK REL=stylesheet HREF="../../global.css" TYPE="text/css">
</head>
<body><h1>Edit Report Narrative</h1>
<form action="<% = Request.ServerVariables("script_name") %>" method="POST">
<INPUT type="hidden" name="inspecID" value="<%= inspecID%>">
<textarea cols="60" rows="20" name="narrative"><%= REPLACE(narrative,"#@#","'") %></textarea><br><br>
<input type="Submit" value="Save Narrative">&nbsp;
<input type="BUTTON" onClick="window.close();" value="Close Window">
</form></body></html>