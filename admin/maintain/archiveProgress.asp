Progress.asp
<% 
'-----------------------------------------------------------------------
'--- This is the progress indicator itself.  It refreshes every second
'--- to re-read the file progress properties, which are updated thoughout
'--- the upload.
'-----------------------------------------------------------------------
'--- Declarations
Dim oFileUpProgress
Dim intProgressID
Dim intPercentComplete
Dim intBytesTransferred
Dim intTotalBytes
Dim bDone

intPercentComplete = 0
intBytesTransferred = 0
intTotalBytes = 0

'--- Instantiate the FileUpProgress object
Set oFileUpProgress = Server.CreateObject("Softartisans.FileUpProgress")

'--- Set the ProgressID with the value we submitted from the form page
oFileUpProgress.ProgressID = CInt(Request.QueryString("progressid"))

'--- Read the values of the progress indicator's properties
intPercentComplete = oFileUpProgress.Percentage
intBytesTransferred = oFileUpProgress.TransferredBytes
intTotalBytes = oFileUpProgress.TotalBytes

%>
<html>
<Head>
<%
	'--- If the upload isn't complete, continue to refresh
	If intPercentComplete < 100 Then
		bDone = False
		Response.Write("<Meta HTTP-EQUIV=""Refresh"" CONTENT=1>")
	Else
		bDone = True
	End If
%>
</head>
<Body>
<TABLE border=1>
<TR>	
<TD colspan=3><B>FileUp Progress Indicator</B></TD>
<TD colspan=2><B>Status: <%If bDone Then Response.Write("Complete!") Else Response.Write("Sending") End If%></B> 
</TR>
<TR><TD>Progress ID </TD>
	<TD>Graphic Indicator</TD>
	<TD>Transferred Bytes</TD>
	<TD>Total Bytes</TD>
	<TD>Transferred Percentage</TD>
</TR>
<TR><TD align=left><%=oFileUpProgress.progressid%></TD>
	<TD>

		<TABLE border=1 cellspacing=0 ALIGN="left" WIDTH="<%=intPercentComplete%>%">
		<TR>
			<TD align=right width="100%" BGCOLOR="blue"><B><%=intPercentComplete%>%</B></TD>
		</TR>
		</TABLE>
<%
	Response.Write("</TD>")
	Response.Write "<TD align=center>" & intBytesTransferred & "</TD>"
	
	if oFileUpProgress.totalbytes > 0 then
		Response.Write("<TD align=center>" & intTotalBytes & "</TD>" & _
		"<TD align=center>" & intPercentComplete & "%</TD>" )
	else
		Response.Write ("<TD align=center>" & "N/A" & "</TD>" & _
		"<TD align=center>" & "N/A" & "</TD>")
	end if
	Response.Write("</TR>")
	
%>
</Table>
</Body>
</Html>
