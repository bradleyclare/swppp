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

If Request.Form.Count > 0 Then
    'insert or update answers to database
    numQuestions = 26
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
<%
'get questions
SQL0 = "SELECT * FROM HortonQuestions ORDER BY orderby"
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)%>
    
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

        If RSA.EOF Then 'if no answers exist start with defaults otherwise use previous answers
            default_val = Trim(RS0("default_answer"))
        Else
            default_val = Trim(RSA("Q" & cnt))
            'response.write(cnt & " - " & default_val & " " )
        End If %>
        
        <tr bgcolor="<%= altColors %>">
        <td><% = question %></td>
        <td>
        <select name="A:<%=cnt%>">
        <option value="yes" <% If default_val = "yes" Then %> selected <% End If %>>yes</option>
        <option value="no" <% If default_val = "no" Then %> selected <% End If %>>no</option>
        <% if (include_na) = True Then %>
            <option value="na" <% If default_val = "na" Then %> selected <% End If %>>n/a</option>
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
    <td></td>
    <td><input name="sync_btn" type="submit" style="font-size: 15px;" value="Sync Answers with Items"/></td>
    <td><input name="submit_btn" type="submit" style="font-size: 15px;" value="Submit"/></td>
    </tr>
    </table>
    </form>
<% End If %>
</BODY>
</HTML>
<% connSWPPP.close
SET connSWPPP=nothing %>