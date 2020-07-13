<%Response.Buffer = False%>
<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and not Session("validAdmin") then Response.Redirect("../default.asp") End If
%><!-- #INCLUDE FILE="../connSWPPP.asp" --><%

inspecID = Session("inspecID")
IF Request("inspecID")<>"" THEN 
	inspecID = Request("inspecID") 
	Session("inspecID")=inspecID
END IF

'get answer data if available
answerSQLSELECT = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
Set RSA = connSWPPP.execute(answerSQLSELECT)

'get questions
SQL0 = "SELECT * FROM HortonQuestions ORDER BY orderby"
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)

numQuestions = 26
If Request.Form.Count > 0 Then
    if Request.Form("sync_btn") = "Sync/Submit" then
        'reset all variables to defaults
        answerSQL = "UPDATE HortonAnswers SET "
        cnt = 0
        Do While Not RS0.EOF
            cnt = cnt + 1
            default_val = Trim(RS0("default_answer"))
            answerSQL = answerSQL & "Q" & cnt & " = '" & default_val & "'"    
            If cnt < numQuestions Then
                answerSQL = answerSQL & ", "
            End If
            RS0.MoveNext
        Loop 'RSO
        answerSQL = answerSQL & " WHERE inspecID = " & inspecID
        'Response.Write(answerSQL)
        connSWPPP.Execute(answerSQL)
        
        'load current items in report and look for categories
        coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
        'Response.Write(coordSQLSELECT)
        Set rsCoord = connSWPPP.execute(coordSQLSELECT)
        repeat_item_found = false
        bmp_issue_found = false
        If rsCoord.EOF Then
            Response.Write("No Items found to sync to!")
        End If
        Do While Not rsCoord.EOF
            'Response.Write("ID: " & rsCoord("coID"))
            if rsCoord("repeat") then
                repeat_item_found = True
            end if
            if rsCoord("pond") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q12 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("outfall") then
                answerSQL = "UPDATE HortonAnswers SET Q13 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("sedloss") then
                answerSQL = "UPDATE HortonAnswers SET Q14 = 'yes' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
			if rsCoord("sedlossw") then
                answerSQL = "UPDATE HortonAnswers SET Q15 = 'yes' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("ce") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q16 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("street") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q17 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("sfeb") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q18 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("rockdam") or rsCoord("riprap") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q19 = 'no' WHERE inspecID = " & inspecID  
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("ip") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q20 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("wo") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q21 = 'no' WHERE inspecID = " & inspecID  
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("veg") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q22 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
			if rsCoord("stock") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q23 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("toilet") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q24 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("trash") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q25 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("dewater") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q26 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            if rsCoord("dust") then
                bmp_issue_found = true
                answerSQL = "UPDATE HortonAnswers SET Q26 = 'no' WHERE inspecID = " & inspecID 
                connSWPPP.Execute(answerSQL)
            end if
            rsCoord.MoveNext
	    Loop 	
        if repeat_item_found then
            answerSQL = "UPDATE HortonAnswers SET Q4 = 'no' WHERE inspecID = " & inspecID  
            connSWPPP.Execute(answerSQL)
        Else
            answerSQL = "UPDATE HortonAnswers SET Q4 = 'yes' WHERE inspecID = " & inspecID  
            connSWPPP.Execute(answerSQL)
        end if
        if bmp_issue_found then
            answerSQL = "UPDATE HortonAnswers SET Q10 = 'no' WHERE inspecID = " & inspecID  
            connSWPPP.Execute(answerSQL)
        Else
            answerSQL = "UPDATE HortonAnswers SET Q10 = 'yes' WHERE inspecID = " & inspecID  
            connSWPPP.Execute(answerSQL)
        end if
	Else
        'insert or update answers to database
        If RSA.EOF Then
            answerSQL = "INSERT INTO HortonAnswers (inspecID, " 
            For i = 1 To numQuestions
                answerSQL = answerSQL & "Q" & i
                If i < numQuestions Then
                    answerSQL = answerSQL & ", "
                End If
            Next
            answerSQL = answerSQL & ") VALUES (" & inspecID & ", "
            For i = 1 To numQuestions
                answerSQL = answerSQL & "'" & TRIM(Request("A:" & i)) & "'"
                If i < numQuestions Then
                    answerSQL = answerSQL & ", "
                End If
            Next
            answerSQL = answerSQL & ")"
        Else
            answerSQL = "UPDATE HortonAnswers SET "
            For i = 1 To numQuestions
                answerSQL = answerSQL & "Q" & i & " = '" & TRIM(Request("A:" & i)) & "'"
                If i < numQuestions Then
                    answerSQL = answerSQL & ", "
                End If
            Next
            answerSQL = answerSQL & " WHERE inspecID = " & inspecID
        End If
        'response.Write(answerSQL)
        connSWPPP.Execute(answerSQL)
    End If
End If

'get updated answer data if changed
answerSQLSELECT = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
Set RSA = connSWPPP.execute(answerSQLSELECT)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: DR Horton Questions</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>
<BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0 marginwidth="5" marginheight="5">
<!-- #INCLUDE FILE="../adminHeader2.inc" -->  

<% RS0.MoveFirst %>

<h1>DR Horton Report Questions</h1>           
<% If RS0.EOF Then %>
	<p>No Questions Found</p>
<% Else %>
    <form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
    <table>
    <tr>
    <th>Question</th>
    <th>Answer</th>
    <th>Category</th>
    <th>Shorthand</th>
    </tr>

    <% cnt = 0 
    altColors="#ffffff"
    Do While Not RS0.EOF
        cnt = cnt + 1
        question = Trim(RS0("question"))
        chkbx_txt = Trim(RS0("chkbx_txt"))
        category = Trim(RS0("category"))
        include_na = RS0("na")
        default_val = Trim(RS0("default_answer"))
        selected_val = ""

        If RSA.EOF Then 'if no answers exist start with defaults otherwise use previous answers
            selected_val = default_val
        Else
            selected_val = Trim(RSA("Q" & cnt))
            'response.write(cnt & " - " & default_val & " " )
        End If %>
        
        <% green = "#6cd97e"
        red = "#e89298" %>
        <tr bgcolor="<%= altColors %>">
        <td><% =cnt %> : <% =question %></td>
        <td>
        <select name="A:<%=cnt%>" <% If default_val = selected_val or selected_val = "na" Then %> style="background-color:<%=green%>;" <% Else %> style="background-color:<%=red%>;"<% End If %>>
        <option value="yes" <% If selected_val = "yes" Then %> selected <% End If %>>yes</option>
        <option value="no" <% If selected_val = "no" Then %> selected <% End If %>>no</option>
        <% if (include_na) = True Then %>
            <option value="na" <% If selected_val = "na" Then %> selected <% End If %>>n/a</option>
	    <% End if %>
        </select>
        </td>
        <td><% = category %></td>
        <td><% = chkbx_txt %></td>
        </tr>
        <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
        RS0.MoveNext
    Loop 'RSO
    RS0.Close
    SET RS0=nothing %>
    <tr><td></td>
    <td><input name="sync_btn" type="submit" style="font-size: 15px;" value="Sync/Submit"/></td>
    <td><input name="submit_btn" type="submit" style="font-size: 15px;" value="Submit"/></td>
    </tr>
    </table>
    </form>
<% End If %>
</BODY>
</HTML>
<% connSWPPP.close
SET connSWPPP=nothing %>