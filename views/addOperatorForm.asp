<%@ Language="VBScript" %>
<!-- #include virtual="admin/connSWPPP.asp" --><% 
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

SQL0="SELECT * FROM ProjectsUsers WHERE userID="& userID &" AND rights='action' AND projectID="& projectID
SET RS0=connSWPPP.execute(SQL0)
IF RS0.eof THEN
	IF NOT(Session("validAdmin") OR Session("validDirector")) THEN
		RS0.Close
		Set RS0=nothing
		connSWPPP.Close
		SET conSWPPP=nothing 
		Response.Write("<SCRIPT language=VBScript>window.close()</SCRIPT>")
		Response.End
	END IF
END IF

If Request.Form.Count > 0 Then

    if inStr(Request("addme"),":") > 0 then
        xArr = Split(Request("addme"),":")
        section = xArr(0)
        action = xArr(1)
        opID = xArr(2)
        select case section
            case "grading"
                param1 = CleanText(Request("gradingDate:"& opID))
                param2 = ""
                nSec = 1
            case "portion"
                param1 = CleanText(Request("portionDate:"& opID))
                param2 = Request("portionText:"& opID)
                nSec = 2
            case "stabilization"
                param1 = CleanText(Request("stabilizationDate:"& opID))
                param2 = CleanText(Request("stabilizationText:"& opID))
                nSec = 3
         End Select 
             SQL2 = "exec usp_OPForm '"& param1 &"', '"& param2 &"', "& opID &", "& action &", "& projectID &", "& userID &", "& nSec
'             response.Write(SQL2)

         If IsDate(param1) Then
             SET RS2=connSWPPP.execute(SQL2)
             RS0.Close
             Set RS0=nothing
         Else
            someError = true
         End If
     end if
End If

function CleanText(textStr)
	CleanText=REPLACE(textStr,"/*","")
	CleanText=REPLACE(CleanText,"*/","")
	CleanText=REPLACE(CleanText,chr(34),"&quot;")
	CleanText=REPLACE(CleanText,chr(39),"&apos;")
	CleanText=REPLACE(CleanText,chr(45),"&hyphen;")
	CleanText=REPLACE(CleanText,chr(145),"&lsquo;")
	CleanText=REPLACE(CleanText,chr(147),"&ldquo;")
	CleanText=REPLACE(CleanText,chr(169),"&copy;")
	CleanText=REPLACE(CleanText,chr(174),"&reg;")
end function

'function addPortion()
'    Form.Item("addme").Value = "portion:2:0"
'end function
%><!-- #include file="cleaner.vb" --><%
SQL1="SELECT o.[OPFormID],o.[projectID],o.[orig_UserID],o.[edit_userID],[edit_userName]= LTrim(RTrim(u.FirstName)) +' '+ LTrim(RTrim(u.LastName)),Convert(char(10),[editDate], 101) as [editDate],[OPFormSection],[OPFormText],[SectionSortby],[SectionSequence], [editable] = Case When (orig_UserID = "& userID &" or edit_UserID = "& userID &") then 'True' Else 'False' End"&_
    " FROM [dbo].[OPForms] o inner join [dbo].[Users] u on o.edit_userID = u.userID WHERE projectID="& projectID &" ORDER BY SectionSortby, SectionSequence" %>
<!-- #include virtual="admin/connSWPPP.asp" --><%
SET RS1=connSWPPP.execute(SQL1)
%>
<html><head>
<title>SWPPP INSPECTIONS - Operator Form</title>
<link rel="stylesheet" type="text/css" href="../global.css"></head>
<body bgcolor="#ffffff" marginwidth="30" leftmargin="30" marginheight="15" topmargin="15">
<center>
<br/><img src="../images/b&wlogoforreport.jpg" width="300"><br><br>
<table cellpadding="2" cellspacing="0" border="0" width="90%">
    <tr align="center"><th><b>OPERATOR FORM</b></th></tr>
</table><br />
<% IF someError THEN %><p style="font-color:red;">There was an error in either the date field or the TextBox.</p><% END IF %>
<FORM action="<% = Request.ServerVariables("script_name") %>" method="post">
<INPUT type="hidden" name="pID" value="<%=projectID%>">
<table cellpadding="2" cellspacing="0" border="0" width="90%">
	<tr><th width="240" align="left">The dates when major grading activities occur:</th>
		<th width="40" align="left"></th>
		<th width="60" align="center">date (mm/dd/yyyy)</th>
		<th width="100" align="center"></th>
	</tr><% 
	If not RS1.Bof and RS1.Eof Then RS1.MoveFirst() End If
	Do While not RS1.Eof
	    If RS1("SectionSortby") = 1 Then %>
    <tr>
        <td><% If RS1("editable") = "True" Then %><input type=radio name="addme" value="grading:1:<%= RS1("OPFormID") %>" />edit&nbsp;<input type=radio name="addme" value="grading:0:<%= RS1("OPFormID") %>" />del<% End If %></td>
        <td><%= RS1("edit_userName") %></td>
        <td><input type="text" size="10" maxlength="10" name="gradingDate:<%= RS1("OPFormID") %>" value="<%= RS1("editDate") %>" <% If RS1("editable") = "False" Then %> readonly<% End If %> /></td>
        <td></td>
      </tr><% 
	    End If
	    RS1.MoveNext()
	Loop %>
    <tr>
        <td><input type=radio name="addme" value="grading:2:0" />add</td>
        <td><%= userName %></td>
        <td><input type="text" size="10" name="gradingDate:0" maxlength="10" value=""/></td>
        <td></td>
    </tr>
	<tr><th width="240" align="left">The dates when construction activities temporarily or permanently cease on a portion of the site:</th>
		<th width="40" align="left"></th>
		<th width="60" align="center" valign="bottom">date (mm/dd/yyyy)</th>
		<th width="100" align="left" valign="bottom">portion of the site</th>
	</tr><% 
	If not RS1.Bof and RS1.Eof Then RS1.MoveFirst() End If
	Do While not RS1.Eof
	    If RS1("SectionSortby") = 2 Then 
	    %>
    <tr>
        <td><% If RS1("editable") = "True" Then %><input type=radio name="addme" value="portion:1:<%= RS1("OPFormID") %>" />edit&nbsp;<input type=radio name="addme" value="portion:0:<%= RS1("OPFormID") %>" />del<% End If %></td>
        <td><%= RS1("edit_userName") %></td>
        <td><input type="text" size="10" name="portionDate:<%= RS1("OPFormID") %>" maxlength="10" value="<%= RS1("editDate") %>" <% If RS1("editable") = "False" Then %> readonly<% End If %> /></td>
        <td><input type="text" maxlength=100 size=25 name="portionText:<%= RS1("OPFormID") %>"<% IF RS1("editable") = "False" Then %>disabled<% END IF %> value="<%= RS1("OPFormText") %>"/></td>
    </tr><% End If
	    RS1.MoveNext()
	Loop %>
    <tr>
        <td><input type=radio name="addme" value="portion:2:0" />add</td>
        <td><%= userName %></td>
        <td><input name="portionDate:0" type="text" size="10" maxlength="10" value="" /></td>
        <td><input type="text" maxlength=100 size=25 name="portionText:0" /></td>
    </tr>
	<tr><th width="240" align="left">The dates when stabilization measures are initiated:</th>
		<th width="40" align="left"></th>
		<th width="60" align="center" valign="bottom">date (mm/dd/yyyy)</th>
		<th width="100" align="left" valign="bottom">stabilization measure</th>
	</tr><% 
	If not RS1.Bof and RS1.Eof Then RS1.MoveFirst() End If
	Do While not RS1.Eof
	    If RS1("SectionSortby") = 3 Then 
	    %><tr><td><% If RS1("editable") = "True" Then %><input type=radio name="addme" value="stabilization:1:<%= RS1("OPFormID") %>" />edit&nbsp;<input type=radio name="addme" value="stabilization:0:<%= RS1("OPFormID") %>" />del<% End If %></td>
	        <td><%= RS1("edit_userName") %></td>
	        <td><input type="text" size="10" maxlength="10" name="stabilizationDate:<%= RS1("OPFormID") %>" value="<%= RS1("editDate") %>" <% If RS1("editable") = "False" Then %> readonly<% End If %> /></td>
	        <% 
	        If RS1("editable") = "True" Then
	         %><td><select name="stabilizationText:<%= RS1("OPFormID") %>">
                    <option<% If RS1("OPFormText") = "construction entrance" Then%> selected<% End If%>>construction entrance</option>
                    <option<% If RS1("OPFormText") = "silt fence" Then%> selected<% End If%>>silt fence</option>
                    <option<% If RS1("OPFormText") = "rock dam" Then%> selected<% End If%>>rock dam</option>
                    <option<% If RS1("OPFormText") = "erosion blanket" Then%> selected<% End If%>>erosion blanket</option>
                    <option<% If RS1("OPFormText") = "riprap" Then%> selected<% End If%>>riprap</option>
                    <option<% If RS1("OPFormText") = "retaining wall" Then%> selected<% End If%>>retaining wall</option>
                    <option<% If RS1("OPFormText") = "slab poured" Then%> selected<% End If%>>slab poured</option>
                    <option<% If RS1("OPFormText") = "paving" Then%> selected<% End If%>>paving</option>
                    <option<% If RS1("OPFormText") = "misc. concrete" Then%> selected<% End If%>>misc. concrete</option>
                    <option<% If RS1("OPFormText") = "hydromulch/seeding/sod" Then%> selected<% End If%>>hydromulch/seeding/sod</option>
                </select></td>
	    </tr><% 
	        Else
	        %><td><%= RS1("OPFormText")%></td><%
	        End If
	    End If
	    RS1.MoveNext()
	Loop 
    %><tr><td><input type=radio name="addme" value="stabilization:2:0" />add</td>
        <td><%= userName %></td><td><input name="stabilizationDate:0" type="text" size="10" maxlength="10" value=""/></td>
        <td><select name="stabilizationText:0">
                <option selected></option>
                <option>construction entrance</option>
                <option>silt fence</option>
                <option>rock dam</option>
                <option>erosion blanket</option>
                <option>riprap</option>
                <option>retaining wall</option>
                <option>slab poured</option>
                <option>paving</option>
                <option>misc. concrete</option>
                <option>hydromulch/seeding/sod</option>
            </select></td>
    </tr>
	<tr><td colspan="4"><br /><br />I certify under penalty of law that this document and all attachments were prepared under my direction or supervision in
	accordance with a system designed to assure that qualified personnel properly gathered and evaluated the information submitted. Based upon
	my inquiry of the person or persons who manage the system, or those persons directly responsible for gathering the information, the information 
	submitted is, to the best of my knowledge and belief, true, accurate, and complete. I am aware that there are significant penalties for 
	submitting false information, including the possibility of fine and imprisonment for knowing voilations.</td></tr>
</table><br><br>
<input type="submit" Value="Edit Operator Form"></FORM></center>
</body></html><%
		RS1.Close
		Set RS1=nothing
		connSWPPP.Close
		SET conSWPPP=nothing  %>