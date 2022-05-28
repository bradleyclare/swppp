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
'Response.Write(answerSQLSELECT)
Set RSA = connSWPPP.execute(answerSQLSELECT)

'get questions
SQL0 = "SELECT * FROM HortonQuestions WHERE orderby > 90 AND orderby < 101 ORDER BY orderby"
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)

inspecSQLSELECT = "SELECT projectID, projectName, projectPhase FROM Inspections WHERE inspecID = " & inspecID
'Response.Write(inspecSQLSELECT)
Set RSI = connSWPPP.execute(inspecSQLSELECT)
projtID = 0
projName = ""
projPhase = ""
If Not RSI.EOF Then
    projID = RSI("projectID")
    projName = Trim(RSI("projectName"))
    projPhase = Trim(RSI("projectPhase"))
End If

numQuestions = 10
If Request.Form.Count > 0 Then
    If Request.Form("default_btn") = "set to defaults" Then
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
    ElseIf Request.Form("previous_btn") = "set to previous report" Then
        'determine previous report id
        If projID > 0 Then
            inspecSQLSELECT = "SELECT inspecID, inspecDate FROM Inspections WHERE projectID = " & projID & " ORDER BY inspecDate DESC"
            'Response.Write(inspecSQLSELECT)
            Set RSII = connSWPPP.execute(inspecSQLSELECT)
            found_current = 0
            prevInspecID = 0
            Do While Not RSII.EOF
                ID = RSII("inspecID")
                'Response.Write("inspecID:" & ID & "</br>")
                If found_current Then
                    prevInspecID = ID
                    'Response.Write("prevInspecID:" & ID & "</br>")
                    Exit Do
                End If
                If Trim(ID) = Trim(inspecID) Then
                    found_current = 1
                    'Response.Write("found_current:" & ID & "=" & inspecID & "</br>")
                End If
                RSII.MoveNext
            Loop

            If prevInspecID > 0 Then
                'get previous horton answers
                answerSQLSELECT = "SELECT * FROM HortonAnswers WHERE inspecID = " & prevInspecID
                'Response.Write(answerSQLSELECT)
                Set RSPA = connSWPPP.execute(answerSQLSELECT)
                If Not RSPA.EOF Then
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
                        if i=3 Then
                            answerSQL = answerSQL & "'yes'"
                            Else
                                answerSQL = answerSQL & "'" & RSPA("Q" & i) & "'"
                            End If
                            If i < numQuestions Then
                                answerSQL = answerSQL & ", "
                            End If
                        Next
                        answerSQL = answerSQL & ")"
                    Else
                        answerSQL = "UPDATE HortonAnswers SET "
                        cnt = 0
                        If Not RSPA.EOF Then
                            For i = 1 To numQuestions
                                current_val = TRIM(RSPA("Q" & i))
                                answerSQL = answerSQL & "Q" & i & " = '" & current_val & "'"
                                If i < numQuestions Then
                                    answerSQL = answerSQL & ", "
                                End If
                            Next
                        End If 'RSPA
                        answerSQL = answerSQL & " WHERE inspecID = " & inspecID
                    End If
                    'Response.Write(answerSQL)
                    connSWPPP.Execute(answerSQL)
                End If

                locationSQLSELECT = "SELECT * FROM HortonLocations WHERE inspecID = " & prevInspecID
                'Response.Write(locationSQLSELECT)
                Set RSL = connSWPPP.execute(locationSQLSELECT)
                If Not RSL.EOF Then
                    startSQL = "INSERT INTO HortonLocations (inspecID, locationName, isOutfall, answer) VALUES"
                    Do While Not RSL.EOF 
                        outfallFlag = 0
                        if RSL("isOutfall") Then
                            outfallFlag = 1
                        End If
                        insertSQL = startSQL &" ("& inspecID &", '"& Trim(RSL("locationName")) &"', "& outfallFlag &", '"& Trim(RSL("answer")) &"')"
                        'Response.Write(insertSQL & "</br>")
                        connSWPPP.Execute(insertSQL)
                        RSL.MoveNext
                    Loop
                End If
            Else
                Response.Write("Could not sync answers to previous report because no previous report found.")
            End If 'prevInspecID
        End If 'end inspections
    ElseIf Request.Form("sync_btn") = "sync" then
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
            'Response.Write(answerSQL & "</br>")
            connSWPPP.Execute(answerSQL)
        Else
            'reset all variables to defaults
            answerSQL = "UPDATE HortonAnswers SET "
            cnt = 0
            Do While Not RS0.EOF
                cnt = cnt + 1
                default_val = Trim(RS0("default_answer"))
                current_val = TRIM(Request("A:" & cnt))
                If current_val <> "na" Then
                    If default_val = "na" Then
                        answerSQL = answerSQL & "Q" & cnt & " = 'yes'" 'if na is the default option but it isn't being used set it to yes
                    Else
                        answerSQL = answerSQL & "Q" & cnt & " = '" & default_val & "'"
                    End IF
                Else
                    answerSQL = answerSQL & "Q" & cnt & " = '" & current_val & "'"
                End If
                If cnt < numQuestions Then
                    answerSQL = answerSQL & ", "
                End If
                RS0.MoveNext
            Loop 'RSO
            answerSQL = answerSQL & " WHERE inspecID = " & inspecID
            'Response.Write(answerSQL & "</br>")
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
                repeat = rsCoord("repeat")
                pond = rsCoord("pond")
                sedloss = rsCoord("sedloss")
                sedlossw = rsCoord("sedlossw")
                ce = rsCoord("ce")
                street = rsCoord("street")
                sfeb = rsCoord("sfeb")
                rockdam = rsCoord("rockdam")
                ip = rsCoord("ip")
                wo = rsCoord("wo")
                veg = rsCoord("veg")
                stock = rsCoord("stock")
                toilet = rsCoord("toilet")
                trash = rsCoord("trash")
                dewater = rsCoord("dewater")
                dust = rsCoord("dust")
                riprap = rsCoord("riprap")
                outfall = rsCoord("outfall")
                intop = rsCoord("intop")
                swalk = rsCoord("swalk")
                mormix = rsCoord("mormix")
                ada = rsCoord("ada")
                dway = rsCoord("dway")
                flume = rsCoord("flume")
                dis = rsCoord("discharge")
                if repeat then
                    repeat_item_found = True
                end if
                if street then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q2 = 'yes' WHERE inspecID = " & inspecID  
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if sfeb or rockdam then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q3 = 'no' WHERE inspecID = " & inspecID  
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if ip or intop then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q4 = 'no' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if mormix or toilet or trash then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q5 = 'no' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if dewater or dust or wo then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q6 = 'no' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if pond then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q7 = 'no' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if veg then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q8 = 'no' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if ada or sedloss or street or swalk or riprap then
                    answerSQL = "UPDATE HortonAnswers SET Q9 = 'yes' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                if dis then
                    answerSQL = "UPDATE HortonAnswers SET Q10 = 'yes' WHERE inspecID = " & inspecID 
                    'Response.Write(answerSQL & "</br>")
                    connSWPPP.Execute(answerSQL)
                end if
                rsCoord.MoveNext
            Loop
        End If
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
'response.Write(answerSQLSELECT)
Set RSA = connSWPPP.execute(answerSQLSELECT)
'If RSA.EOF Then Response.Write("<p>no answers found</p>") End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Forestar Questions</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>
<BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0 marginwidth="5" marginheight="5">
<!-- #INCLUDE FILE="../adminHeader2.inc" -->  

<% RS0.MoveFirst %>

<h1>Forestar report questions</h1>           
<% If RS0.EOF Then %>
	<p>no questions found</p>
<% Else %>
    <form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
    <table>
    <tr>
    <th>question</th>
    <th>answer</th>
    <th>category</th>
    <th>shorthand</th>
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
            'response.write(cnt & " - s:" & selected_val & " - d:" & default_val & " - ina:" & include_na & "</br>" )
        End If %>
        
        <% green = "#6cd97e"
        red = "#e89298" %>
        <tr bgcolor="<%= altColors %>">
        <td><% =cnt %> : <% =question %></td>
        <td>

        <% show_dropdown = 1
        If show_dropdown Then 
            If selected_val = "na" and include_na = False Then 
                selected_val = default_val
            End If 
            If default_val = "na" Then
                default_val = "yes"
            End If 
            %>
            <select name="A:<%=cnt%>" 
            <% If default_val = selected_val or selected_val = "na" Then %> 
                style="background-color:<%=green%>;" 
            <% Else %> 
                style="background-color:<%=red%>;"
            <% End If %>
            >
            <option value="yes" <% If selected_val = "yes" Then %> selected <% End If %>>yes</option>
            <option value="no" <% If selected_val = "no" Then %> selected <% End If %>>no</option>
            <% if (include_na) = True Then %>
                <option value="na" <% If selected_val = "na" Then %> selected <% End If %>>n/a</option>
	        <% End if %>
            </select>
        <% End If %>
        
        </td>
        <td><% = category %></td>
        <td><% = chkbx_txt %></td>
        </tr>

        <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
        RS0.MoveNext
    Loop 'RSO
    RS0.Close
    SET RS0=nothing %>
    <tr><td>
    <input name="default_btn" type="submit" style="font-size: 15px;" value="set to defaults"/>
    <input name="previous_btn" type="submit" style="font-size: 15px;" value="set to previous report"/>
    </td>
    <td>
    <% If Not RSA.EOF Then %>
    <input name="sync_btn" type="submit" style="font-size: 15px;" value="sync"/>
    <% End If %>
    </td>
    <td><input name="submit_btn" type="submit" style="font-size: 15px;" value="submit"/></td>
    </tr>
    </table>
    </form>
<% End If %>
</BODY>
</HTML>
<% connSWPPP.close
SET connSWPPP=nothing %>