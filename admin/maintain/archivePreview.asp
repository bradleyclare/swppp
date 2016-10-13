<%@ Language="VBScript" %>
<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If

function validated(xValue, xType)
	validated=True
	SELECT CASE xType
		CASE "directory"
			SET re= NEW RegExp
			re.Global=true
			re.IgnoreCase=true
			re.pattern="([a-zA-Z]:\\([\w ]*)*\\)|([a-zA-Z]:\\)"
			IF NOT(re.Test(xValue)) THEN validated=False END IF
		CASE "integer"
			IF NOT(IsNumeric(xValue)) THEN validated=False END IF
	END SELECT
end function
'--IF NOT(validated(Request("destination"),"directory")) THEN Response.Redirect("archive.asp?err=001") END IF
IF NOT(validated(Request("projectID"),"integer")) THEN Response.Redirect("archive.asp?err=002") END IF

SQL0="sp_GetAllInspectionsforProject("& Request("projectID") &")"
%><!-- #include file="../connSWPPP.asp" --><%
SET RS0=connSWPPP.Execute(SQL0)
cnt0=0
DO WHILE NOT RS0.EOF
	cnt0=cnt0+1
	RS0.MoveNext
LOOP
RS0.MoveFirst %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : Maintain : Archive Preview</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
<script language="JavaScript" src="../js/validUpload.js"></script>
<script language="JavaScript" src="../js/validUpload1.2.js"></script>
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td><h1>SWPPP Inspections : Maintain : Archive Preview</h1></td>
	</tr>
</table>
<FORM action="archiveCreateFileStructure.asp" method="post" name="theForm">
<input type="hidden" name="projectID" value="<%= Request("projectID")%>">
<input type="hidden" name="destination" value="<%= Request("destination")%>">
<input type="submit" value="Create File Structure">
</FORM>

This page will display the file structure for the final archival.<br>
<TABLE border="1">
	<TR><TD><TABLE border="0" align="left" cellpadding=0 cellspacing=0 margin=0>
			<TR><TD align="right" class="filetree"><IMG src="..\..\images\FileTreeTopFolder.gif"></TD>
				<TD colspan="3">Root Folder</TD></TR>
			<TR><TD align="right" class="filetree"><IMG src="..\..\images\FileTreeT.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeText.gif"></TD>
				<TD colspan="2">autorun.inf</TD></TR>
			<TR><TD align="right" class="filetree"><IMG src="..\..\images\FileTreeT.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeText.gif"></TD>
				<TD colspan="2">Reports.asp</TD></TR>
			<TR><TD align="right" class="filetree"><IMG src="..\..\images\FileTreeT.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeFolder.gif"></TD>
				<TD colspan="2">Images</TD></TR>
			<TR><TD align="right" class="filetree"><IMG src="..\..\images\FileTreeI.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeT.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeJpeg.gif"></TD>
				<TD>logo.jpg</TD></TR>
			<TR><TD align="right" class="filetree"><IMG src="..\..\images\FileTreeI.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeL.gif"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeJpeg.gif"></TD>
				<TD>image.jpg</TD></TR>
<%	DO WHILE NOT RS0.EOF 
		SQL1="sp_GetInspectionData("& RS0(0) &")" 
		SET RS1=connSWPPP.execute(SQL1)
		cnt0=cnt0-1		
		SQL2="sp_GetOptionalImages("& RS0(0) &")"
		SET RS2=connSWPPP.execute(SQL2)
		IF cnt0>0 THEN imgSrc1="..\..\images\FileTreeT.gif" ELSE imgSrc1="..\..\images\FileTreeL.gif" END IF 
		dirName = MonthName(Month(RS0(1))) &"_"& Day(RS0(1)) &"_"& Year(RS0(1)) %>
			<TR><TD align="right" class="filetree"><IMG src="<%=imgSrc1%>"></TD>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeFolder.gif"></TD>
				<TD colspan="2"><%=dirName%></TD></TR>
			<TR><TD align="right" class="filetree"><%IF cnt0>0 THEN%><IMG src="..\..\images\FileTreeI.gif"><%END IF%></TD>
<%			IF RS1.EOF THEN %>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeL.gif"></TD>
<%			ELSE %>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeT.gif"></TD>
<% 			END IF %>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeWebpage.gif"></TD>
				<TD colspan="2">Inspection Report</TD></TR>
<%		DO WHILE NOT RS2.EOF 
			imgTypeDesc=RS2("oitDesc")
			RS2.MoveNext %>
			<TR><TD align="right" class="filetree"><%IF cnt0>0 THEN%><IMG src="..\..\images\FileTreeI.gif"><%END IF%></TD>
<%			IF RS2.EOF THEN %>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeL.gif"></TD>
<%			ELSE %>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeT.gif"></TD>
<% 			END IF %>
				<TD align="right" class="filetree"><IMG src="..\..\images\FileTreeJpeg.gif"></TD>
				<TD><%= imgTypeDesc %></TD></TR><%
		LOOP
		RS0.MoveNext
	LOOP %>						
			</TABLE>	
			</TD></TR>	
		</TD>
	</TR>
</TABLE>
</body>
</html>
<% 	SET RS0=nothing
	SET RS1=nothing
connSWPPP.close() %>