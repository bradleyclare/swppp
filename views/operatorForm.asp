<%@ Language="VBScript" %>
<!-- #include file="../admin/connSWPPP.asp" --><% 
If 	Not Session("validAdmin") And _
	Not Session("validDirector") And _
	Not Session("validInspector") And _
	Not Session("validUser") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") &"?"& Request.ServerVariables("query_string")
	Response.Redirect("../admin/maintain/loginUser.asp")	
End If
projectID = Request("pID")
userID = Session("userID")
userName = Session("FirstName") +" "+ Session("LastName")

%><!-- #include file="cleaner.vb" --><%
SQL1="SELECT o.[OPFormID],o.[projectID],o.[orig_UserID],o.[edit_userID],[edit_userName]= LTrim(RTrim(u.FirstName)) +' '+ LTrim(RTrim(u.LastName)),Convert(char(10),[editDate], 101) as [editDate],[OPFormSection],[OPFormText],[SectionSortby],[SectionSequence], [editable] = Case When (orig_UserID = "& userID &" or edit_UserID = "& userID &") then 'True' Else 'False' End"&_
    " FROM [dbo].[OPForms] o inner join [dbo].[Users] u on o.edit_userID = u.userID WHERE projectID="& projectID &" ORDER BY SectionSortby, SectionSequence" %>
<!-- #include file="../admin/connSWPPP.asp" --><%
SET RS1=connSWPPP.execute(SQL1)
%>
<html><head>
<title>SWPPP INSPECTIONS - Operator Form</title>
<link rel="stylesheet" type="text/css" href="../global.css"></head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
<table cellpadding="2" cellspacing="0" border="0" width="90%">
    <tr align="center"><th><b>OPERATOR FORM</b></th></tr>
</table><br />
<% IF someError THEN %><p style="font-color:red;">There was an error in either the date field or the TextBox.</p><% END IF %>
<FORM action="<% = Request.ServerVariables("script_name") %>" method="post">
<INPUT type="hidden" name="pID" value="<%=projectID%>">
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="240" align="left">The dates when major grading activities occur:</th>
		<th width="40" align="left"></th>
		<th width="60" align="center">date </th>
		<th width="100" align="center"></th>
	</tr><% 
	If not RS1.Bof and RS1.Eof Then RS1.MoveFirst() End If
	Do While not RS1.Eof
	    If RS1("SectionSortby") = 1 Then %>
    <tr>
        <td></td>
        <td><%= RS1("edit_userName") %></td>
        <td><%= RS1("editDate") %></td>
        <td></td>
      </tr><% 
	    End If
	    RS1.MoveNext()
	Loop %>
	<tr><th width="240" align="left">The dates when construction activities temporarily or permanently cease on a portion of the site:</th>
		<th width="40" align="left"></th>
		<th width="60" align="center" valign="bottom">date </th>
		<th width="100" align="left" valign="bottom">portion of the site</th>
	</tr><% 
	If not RS1.Bof and RS1.Eof Then RS1.MoveFirst() End If
	Do While not RS1.Eof
	    If RS1("SectionSortby") = 2 Then 
	    %>
    <tr><td></td><td><%= RS1("edit_userName") %></td><td><%= RS1("editDate") %></td><td><%= RS1("OPFormText") %></td></tr><% 
        End If
	    RS1.MoveNext()
	Loop %>
	<tr><th width="240" align="left">The dates when stabilization measures are initiated:</th>
		<th width="40" align="left"></th>
		<th width="60" align="center" valign="bottom">date </th>
		<th width="100" align="left" valign="bottom">stabilization measure</th>
	</tr><% 
	If not RS1.Bof and RS1.Eof Then RS1.MoveFirst() End If
	Do While not RS1.Eof
	    If RS1("SectionSortby") = 3 Then 
	    %><tr><td></td><td><%= RS1("edit_userName") %></td><td><%= RS1("editDate") %></td><td><%= RS1("OPFormText") %></td></tr><%
	    End If
	    RS1.MoveNext()
	Loop 
    %>
	<tr><td colspan="4"><br /><br />I certify under penalty of law that this document and all attachments were prepared under my direction or supervision in
	accordance with a system designed to assure that qualified personnel properly gathered and evaluated the information submitted. Based upon
	my inquiry of the person or persons who manage the system, or those persons directly responsible for gathering the information, the information 
	submitted is, to the best of my knowledge and belief, true, accurate, and complete. I am aware that there are significant penalties for 
	submitting false information, including the possibility of fine and imprisonment for knowing voilations.</td></tr>
</table><br><br></FORM></center>
</body></html><%
		RS1.Close
		Set RS1=nothing
		connSWPPP.Close
		SET conSWPPP=nothing  %>