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
SQL0 = "SELECT * FROM HortonQuestions WHERE orderby > 60 AND orderby < 87 ORDER BY orderby"
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)

numQuestions = 26
If Request.Form.Count > 0 Then
    If Request.Form("na_btn") = "set to n/a" Then
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
                answerSQL = answerSQL & "'na'"
                If i < numQuestions Then
                    answerSQL = answerSQL & ", "
                End If
            Next
            answerSQL = answerSQL & ")"
        Else
            answerSQL = "UPDATE HortonAnswers SET "
            For i = 1 To numQuestions
                answerSQL = answerSQL & "Q" & i & " = 'na'"
                If i < numQuestions Then
                    answerSQL = answerSQL & ", "
                End If
            Next
            answerSQL = answerSQL & " WHERE inspecID = " & inspecID
        End If
        'response.Write(answerSQL)
        connSWPPP.Execute(answerSQL)
    ElseIf Request.Form("default_btn") = "set to defaults" Then
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
        inspecSQLSELECT = "SELECT projectID FROM Inspections WHERE inspecID = " & inspecID
        'Response.Write(inspecSQLSELECT)
        Set RSI = connSWPPP.execute(inspecSQLSELECT)
        If Not RSI.EOF Then
            projectID = RSI("projectID")
            inspecSQLSELECT = "SELECT inspecID, inspecDate FROM Inspections WHERE projectID = " & projectID & " ORDER BY inspecDate DESC"
            'Response.Write(inspecSQLSELECT)
            Set RSII = connSWPPP.execute(inspecSQLSELECT)
            found_current = 0
            prevInspecID = 0
            Do While Not RSII.EOF
                ID = RSII("inspecID")
                'Response.Write("inspecID:" & ID & "</br>")
                If found_current Then
                    prevInspecID = ID
                    Response.Write("prevInspecID:" & ID & "</br>")
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
                numQuestions = 26
                If Not RSPA.EOF Then
                    If Not RSA.EOF Then                    
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
                        'Response.Write(answerSQL)
                        connSWPPP.Execute(answerSQL)
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
                        'Response.Write(answerSQL)
                        connSWPPP.Execute(answerSQL)
                    End If
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
        Else
            'reset all variables to defaults
            answerSQL = "UPDATE HortonAnswers SET "
            cnt = 0
            Do While Not RS0.EOF
                cnt = cnt + 1
                default_val = Trim(RS0("default_answer"))
                current_val = TRIM(Request("A:" & cnt))
                If cnt = 4 or cnt = 10 or cnt >= 12 Then
                    If current_val <> "na" Then
                        answerSQL = answerSQL & "Q" & cnt & " = '" & default_val & "'"
                    Else
                        answerSQL = answerSQL & "Q" & cnt & " = '" & current_val & "'"
                    End If
                Else
                    answerSQL = answerSQL & "Q" & cnt & " = '" & current_val & "'"
                End If
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
                if repeat then
                    repeat_item_found = True
                end if
                if pond then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q12 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if sedloss then
                    answerSQL = "UPDATE HortonAnswers SET Q14 = 'yes' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if sedlossw then
                    answerSQL = "UPDATE HortonAnswers SET Q15 = 'yes' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if ce then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q16 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if street or intop or swalk or ada or dway or flume then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q17 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if sfeb then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q18 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if rockdam or riprap then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q19 = 'no' WHERE inspecID = " & inspecID  
                    connSWPPP.Execute(answerSQL)
                end if
                if ip then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q20 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if wo or mormix then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q21 = 'no' WHERE inspecID = " & inspecID  
                    connSWPPP.Execute(answerSQL)
                end if
                if veg then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q22 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if stock then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q23 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if toilet then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q24 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if trash then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q25 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if dewater then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q26 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if dust then
                    bmp_issue_found = true
                    answerSQL = "UPDATE HortonAnswers SET Q26 = 'no' WHERE inspecID = " & inspecID 
                    connSWPPP.Execute(answerSQL)
                end if
                if outfall then
                    answerSQL = "UPDATE HortonAnswers SET Q13 = 'no' WHERE inspecID = " & inspecID 
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
            End if
            if bmp_issue_found then
                answerSQL = "UPDATE HortonAnswers SET Q10 = 'no' WHERE inspecID = " & inspecID  
                connSWPPP.Execute(answerSQL)
            Else
                answerSQL = "UPDATE HortonAnswers SET Q10 = 'yes' WHERE inspecID = " & inspecID  
                connSWPPP.Execute(answerSQL)
            End if
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

        'find and update location answers
        maxNumLocations = 20
        finalAnswer = "yes"
        For i = 1 To maxNumLocations
            'response.Write("pond:"& i &"</br>")
            if Trim(Request("pondLocA:" & CStr(i))) = "" then
		        exit for
		    end if
            if Trim(Request("pondLocA:" & CStr(i))) = "no" then
                finalAnswer = "no"
            end if
            answerSQL = "UPDATE HortonLocations SET answer = '"& TRIM(Request("pondLocA:" & CStr(i))) &"' WHERE inspecID = "& inspecID &" AND locationID = "&  TRIM(Request("pondLocID:" & CStr(i)))
            'response.Write(answerSQL)
            connSWPPP.Execute(answerSQL)
        Next
        If i > 1 Then 'update the overall answer if any locations were defined
            answerSQL = "UPDATE HortonAnswers SET Q12 = '"& finalAnswer &"' WHERE inspecID ="& inspecID
            'response.Write(answerSQL)
            connSWPPP.Execute(answerSQL)
        End If

        finalAnswer = "yes"
        For i = 1 To maxNumLocations
            'response.Write("outfall:"& i &"</br>")
            if Trim(Request("outfallLocA:" & CStr(i))) = "" then
		        exit for
		    end if
            if Trim(Request("outfallLocA:" & CStr(i))) = "no" then
                finalAnswer = "no"
            end if
            answerSQL = "UPDATE HortonLocations SET answer = '"& TRIM(Request("outfallLocA:" & CStr(i))) &"' WHERE inspecID = "& inspecID &" AND locationID = "&  TRIM(Request("outfallLocID:" & CStr(i)))
            'response.Write(answerSQL)
            connSWPPP.Execute(answerSQL)
        Next
        If i > 1 Then 'update the overall answer if any locations were defined
            answerSQL = "UPDATE HortonAnswers SET Q13 = '"& finalAnswer &"' WHERE inspecID ="& inspecID
            'response.Write(answerSQL)
            connSWPPP.Execute(answerSQL)
        End If
    End If
End If

'get updated answer data if changed
answerSQLSELECT = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
Set RSA = connSWPPP.execute(answerSQLSELECT)

pondSQL="SELECT * FROM HortonLocations WHERE inspecID="& inspecID &" AND isOutfall=0"
'response.Write(pondSQL)
Set RSpond=connSWPPP.execute(pondSQL)

outfallSQL="SELECT * FROM HortonLocations WHERE inspecID="& inspecID &" AND isOutfall=1"
'response.Write(outfallSQL)
Set RSoutfall=connSWPPP.execute(outfallSQL)
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

<h1>DR Horton report questions</h1>           
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
            'response.write(cnt & " - " & default_val & " " )
        End If %>
        
        <% green = "#6cd97e"
        red = "#e89298" %>
        <tr bgcolor="<%= altColors %>">
        <td><% =cnt %> : <% =question %></td>
        <td>

        <% show_dropdown = 1
        If cnt = 12 Then %>
            <a href="defineHortonLocations.asp?inspecID=<%=inspecID%>"><input type="button" value="define locations"/></a>
            <% If Not RSpond.EOF Then
                show_dropdown = 0
            End If 
        ElseIf cnt = 13 Then %>
            <a href="defineHortonLocations.asp?inspecID=<%=inspecID%>&outfall=1"><input type="button" value="define locations"/></a>
            <% If Not RSoutfall.EOF Then
                show_dropdown = 0
            End If
        End If
        If show_dropdown Then %>
            <select name="A:<%=cnt%>" <% If default_val = selected_val or selected_val = "na" Then %> style="background-color:<%=green%>;" <% Else %> style="background-color:<%=red%>;"<% End If %>>
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
        
        <% 'for question 12 and 13 show the defined location questions 
        If cnt = 12 Then
            pondCnt = 0
            anyNo = 0
            default_val = "yes"
            Do While Not RSpond.EOF
                pondCnt = pondCnt + 1
                locationID = RSpond("locationID")
                locationName = Trim(RSpond("locationName")) 
                selected_val = Trim(RSpond("answer")) %>
                <tr bgcolor="<%= altColors %>">
                <td>&nbsp-&nbsp<% =locationName %></td>
                <td>
                <input type="hidden" name="pondLocID:<%=pondCnt%>" value="<%=locationID%>"/>
                <select name="pondLocA:<%=pondCnt%>" <% If default_val = selected_val or selected_val = "na" Then %> style="background-color:<%=green%>;" <% Else %> style="background-color:<%=red%>;"<% End If %>>
                <option value="yes" <% If selected_val = "yes" Then %> selected <% End If %>>yes</option>
                <option value="no" <% If selected_val = "no" Then %> selected <% End If %>>no</option>
                <option value="na" <% If selected_val = "na" Then %> selected <% End If %>>n/a</option>
                <% If selected_val = "no" Then
                    anyNo = 1
                End If %>
                </select>
                </td>
                <td></td>
                <td></td>
                </tr>
                <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
                RSpond.MoveNext
            Loop
            'set overall answer based on the location answers
            If pondCnt > 0 Then
                If anyNo Then %>
                    <input type="hidden" name="A:<%=cnt%>" value="no"/>
                <% Else %>
                    <input type="hidden" name="A:<%=cnt%>" value="yes"/>
                <% End If
            End If
        End If

        If cnt = 13 Then
            outfallCnt = 0
            anyNo = 0
            default_val = "yes"
            Do While Not RSoutfall.EOF
                outfallCnt = outfallCnt + 1
                locationID = RSoutfall("locationID")
                locationName = Trim(RSoutfall("locationName")) 
                selected_val = Trim(RSoutfall("answer")) %>
                <tr bgcolor="<%= altColors %>">
                <td>&nbsp-&nbsp<% =locationName %></td>
                <td>
                <input type="hidden" name="outfallLocID:<%=outfallCnt%>" value="<%=locationID%>"/>
                <select name="outfallLocA:<%=outfallCnt%>" <% If default_val = selected_val or selected_val = "na" Then %> style="background-color:<%=green%>;" <% Else %> style="background-color:<%=red%>;"<% End If %>>
                <option value="yes" <% If selected_val = "yes" Then %> selected <% End If %>>yes</option>
                <option value="no" <% If selected_val = "no" Then %> selected <% End If %>>no</option>
                <option value="na" <% If selected_val = "na" Then %> selected <% End If %>>n/a</option>
                <% If selected_val = "no" Then
                    anyNo = 1
                End If %>
                </select>
                </td>
                <td></td>
                <td></td>
                </tr>
                <% If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
                RSoutfall.MoveNext
            Loop
            'set overall answer based on the location answers
            If outfallCnt > 0 Then
                If anyNo Then %>
                    <input type="hidden" name="A:<%=cnt%>" value="no"/>
                <% Else %>
                    <input type="hidden" name="A:<%=cnt%>" value="yes"/>
                <% End If
            End If
        End If

        If altColors = "#e5e6e8" Then altColors = "#ffffff" Else altColors = "#e5e6e8" End If
        RS0.MoveNext
    Loop 'RSO
    RS0.Close
    SET RS0=nothing %>
    <tr><td>
    <input name="na_btn" type="submit" style="font-size: 15px;" value="set to n/a"/>
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