<%@ Language="VBScript" %>
<%
testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
If not(Session("validAdmin") AND InStr(testStr,Session("email"))>0) Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("../default.asp")
End If
%><!-- #include file="../connSWPPP.asp" --><%
projectID=Request("id")

IF (IsNull(projectID) OR NOT(IsNumeric(projectID))) THEN Response.redirect("viewProjects.asp") END IF

If Request.Form.Count>0 THEN
	if Request.Form("submit_delete") = "delete" then
		Response.Redirect("deleteProject.asp?id=" +  projectID)
	End If
	if Request.Form("submit_rights") = "rights" then
		Response.Redirect("editUsersByProject.asp?pID=" + projectID)
	End If
	err=0
	active=Request("active")
	projOpenItemAlert = Request("projOpenItemAlert")
	initInspecCost=Request("initInspecCost")
	IF NOT(IsNumeric(initInspecCost)) THEN err=err+4 END IF
	inspecCost=Request("inspecCost")
	IF NOT(IsNumeric(inspecCost)) THEN err=err+8 END IF
	invoiceMemo=Trim(Request("invoiceMemo"))
	IF (Request("bCycle")<1 OR Request("bCycle")>4) THEN err=err+16 END IF
	if err=0 then
		SQL1="UPDATE Projects SET active='"& active &"', projOpenItemAlert='"& projOpenItemAlert &"', initInspecCost="& initInspecCost &", inspecCost="& inspecCost &", invoiceMemo = '"& invoiceMemo &"', billCycle='"& Trim(Request("bCycle")) &"', collectionName = '"&Trim(Request("groupName")) & "'"&_
			" WHERE projectID="& projectID
		'Response.Write(SQL1)
		connSWPPP.Execute(SQL1)	
	else
		err=DecToBin(err)
	end if
End IF
SQL0="SELECT * FROM Projects WHERE projectID="& projectID
SET RS0=connSWPPP.Execute(SQL0) 

SQL1="SELECT * FROM Inspections WHERE projectID="& projectID
'Response.Write(SQL1)
SET RS1=connSWPPP.Execute(SQL1)
DO WHILE NOT RS1.EOF
    Response.Write(RS1("groupName"))
    RS1.MoveNext
LOOP

function validStr(testStr)
	strPassed=True
	for i = 0 to Len(testStr) Step 1
		if (ASC(MID(testStr,i,1))>32 AND ASC(MID(testStr,i,1))<48) then
			strPassed=False
			Exit For
		end if
	next
	validStr=strPassed
end function

Function DecToBin(intDec)
  dim strResult
  dim intValue
  dim intExp

  strResult = ""

  intValue = intDEC
  intExp = 65536
  while intExp >= 1
    if intValue >= intExp then
      intValue = intValue - intExp
      strResult = strResult & "1"
    else
      strResult = strResult & "0"
    end if
    intExp = intExp / 2
  wend

  DecToBin = strResult
End Function

Function BinToDec(strBin)
  dim lngResult
  dim intIndex

  lngResult = 0
  for intIndex = len(strBin) to 1 step -1
    strDigit = mid(strBin, intIndex, 1)
    select case strDigit
      case "0"
        ' do nothing
      case "1"
        lngResult = lngResult + (2 ^ (len(strBin)-intIndex))
      case else
        ' invalid binary digit, so the whole thing is invalid
        lngResult = 0
        intIndex = 0 ' stop the loop
    end select
  next

  BinToDec = lngResult
End Function
%>
<html>
<head>
	<title>SWPPP INSPECTIONS : Admin : Edit Project Info</title>
	<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->

<h1>Edit Project Information</h1>
<form action="<%= Request.ServerVariables("script_name") %>" method="post">
<table width="90%" border="0">
	<input type="hidden" name="id" value="<%= projectID %>">
<% 	IF LEN(err)>0 THEN 
		IF MID(err,Len(err),1)="1" THEN %><tr><td colspan="2"><font color="red">*The Project Name that you entered contains illegal characters*</font></td></tr><% END IF %>
<%	END IF %>
	<tr><th width="30%">Project Name</th>
		<td align=left width="70%"><%= Trim(RS0("projectName"))%></td></tr>
<% 	IF LEN(err)>1 THEN
		IF MID(err,Len(err)-1,1)="1" THEN %><tr><td colspan="2"><font color="red">*The Project Phase that you entered contains illegal characters*</font></td></tr><% END IF %>
<%	END IF %>
	<tr><th>Project Phase</th>
		<td align=left><%= Trim(RS0("projectPhase"))%></td></tr>
	<tr><th>Active Project</th>
		<td align=left><INPUT maxlength="1" name="active"  value="<%= Trim(RS0("active"))%>"></td></tr>
	<tr><th>Open Item Alert</th>
		<td align=left><INPUT maxlength="1" name="projOpenItemAlert"  value="<%= Trim(RS0("projOpenItemAlert"))%>"></td></tr>
<% 	IF LEN(err)>2 THEN
		IF MID(err,Len(err)-2,1)="1" THEN %><tr><td colspan="2"><font color="red">*The Initial Inspection Cost must be a number*</font></td></tr><% END IF %>
<%	END IF %>
	<tr><th>Initial Inspection Cost</th>
		<td align=left><INPUT name="initInspecCost" value="<%= FormatNumber(RS0("initInspecCost"),2)%>"></td></tr>
<% 	IF LEN(err)>3 THEN
		IF MID(err,Len(err)-3,1)="1" THEN %><tr><td colspan="2"><font color="red">*The Recurring Inspection Cost must be a number*</font></td></tr><% END IF %>
<%	END IF %>
	<tr><th>Recurring Inspection Cost</th>
		<td align=left><INPUT name="inspecCost" value="<%= FormatNumber(RS0("inspecCost"),2)%>"></td></tr>
	<tr><th>Invoice Memo</th>
	    <td align=left><input width="200" name="invoiceMemo" value="<%= Trim(RS0("invoiceMemo")) %>" /></td></tr>
	<tr><th>Billing Cycle</th>
		<td align=left><SELECT name="bCycle">
			<OPTION value="1"<% IF RS0("billCycle")=1 THEN %> selected<% END IF %>>1</option>
			<OPTION value="2"<% IF RS0("billCycle")=2 THEN %> selected<% END IF %>>2</option>
			<OPTION value="3"<% IF RS0("billCycle")=3 THEN %> selected<% END IF %>>3</option>
			<OPTION value="4"<% IF RS0("billCycle")=4 THEN %> selected<% END IF %>>4</option>
			</SELECT></td></tr>
    <tr><th>Project Group: </th>
	<td align=left>
        <select name="groupName">
            <option value=""></option>
        <% SQLGroups="SELECT * FROM Groups ORDER BY groupName"
	    SET RSGroups=connSWPPP.execute(SQLGroups)
	    DO WHILE NOT RSGroups.EOF %>	
            <option value="<%= RSGroups("groupName")%>"
            <% If Trim(RS0("collectionName")) = Trim(RSGroups("groupName")) Then %> selected
            <% End If %>><%= Trim(RSGroups("groupName"))%></option>
        <% RSGroups.MoveNext
	    LOOP %>	
        </select>
	</td></tr>
	<tr><td></td></tr>
</table>
<input name="submit_delete" type="submit" value="delete"/>
<input name="submit_rights" type="submit" value="rights"/>
<input type="submit" value="update">
</form>
</body>
</html><%
RS0.Close
Set RS0=nothing
connSWPPP.Close
Set connSWPPP=nothing %>