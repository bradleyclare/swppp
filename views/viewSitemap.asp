<%@ Language="VBScript" %>
<%
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
inspecID = Request("inspecID")
%>
<!-- #include file="../admin/connSWPPP.asp" -->
<% dirName="sitemap"
fileDesc= "Site Map"
SQLa="sp_oImagesByType "& inspecID &",'12'"
SET RSa=connSWPPP.Execute(SQLa)
cnt1=1
curOITDesc=""
DO WHILE NOT(RSa.EOF)
	thisFileDesc=fileDesc
	if curOITDesc=fileDesc then
		cnt1=cnt1+1
	else
		cnt1=1
		curOITDesc=fileDesc
	end if
	if (cnt1>1) then thisFileDesc=thisFileDesc &" "& cnt1 end if
	if dirName = "sitemap" then
		sitemap_name = "../images/"& dirName &"/"& trim(RSa("oImageFileName"))
		Response.Redirect(sitemap_name)
		exit do
	end if
	RSa.MoveNext
LOOP 
Response.Write("No sitemap was found for this project.  If you beleive this is an error please report it to SWPPP.com.<br/>")%>           
