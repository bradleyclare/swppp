<%Response.Buffer = False%>
<%
'Response.Write(Response.Buffer)
' Send Menu Email
' smp 3/5/03 layout
If Not Session("validInspector") and not Session("validAdmin") then Response.Redirect("../default.asp") End If
%><!-- #INCLUDE FILE="../connSWPPP.asp" --><%

Server.ScriptTimeout=1500

'Response.Write(Request.Form.Count & "<br>")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
	<TITLE>SWPPP INSPECTIONS :: Admin :: Cleanup Scoring</TITLE>
	<LINK REL=stylesheet HREF="../../global.css" type="text/css">
</HEAD>
<BODY vLink=#d1a430 aLink=#000000 link=#b83a43 bgColor=#ffffff leftMargin=0 topMargin=0 marginwidth="5" marginheight="5">
<!-- #INCLUDE FILE="../adminHeader2.inc" -->  
<%
'get all projects
SQL0 = "SELECT inspecID, inspecDate, reportType," & _
	" projectID, projectName, projectPhase, released, includeItems, compliance, totalItems, completedItems" & _
	" FROM Inspections" & _
	" WHERE includeItems = 1" &_
    " AND compliance = 0" &_
    " AND openItemAlert = 1" 
'Response.Write(SQL0)
Set RS0 = connSWPPP.Execute(SQL0)%>
    
<h1>Cleanup Scoring</h1>                    
<% If RS0.EOF Then %>
	<p>No Reports Found</p>
<% Else
    fix_db = true
    inspecLimit = 10000
    inspecStart = 1
    inspecEnd = inspecStart + inspecLimit
    inspecCnt = 0
    updateCnt = 0
    Do While Not RS0.EOF
        inspecCnt = inspecCnt + 1
        if inspecCnt < inspecStart then
            'do nothing
        elseif inspecCnt > inspecEnd then
            'do nothing
        else
            projName = Trim(RS0("projectName"))
            projPhase = Trim(RS0("projectPhase"))
            inspecID = RS0("inspecID")
            inspecDate = RS0("inspecDate")
            totalItems = RS0("totalItems")
            completedItems = RS0("completedItems")%>
                    
            <%'open items on report tally up the open item dates 
            coordSQLSELECT = "SELECT coID, status, repeat, infoOnly, LD, parentID FROM Coordinates" &_
	            " WHERE inspecID=" & inspecID &_
                " AND infoOnly=0" &_
                " ORDER BY orderby"	
            'Response.Write(coordSQLSELECT)
            Set rsCoord = connSWPPP.execute(coordSQLSELECT)

            coordCnt = 0
            completedItem_cnt = 0
            repeatItem_cnt = 0
            LDItem_cnt = 0
            If rsCoord.EOF Then
                'do nothing    
	        Else
                Do While Not rsCoord.EOF
                    coordCnt = coordCnt + 1
                    coID = rsCoord("coID")
			        status = rsCoord("status")
			        repeat = rsCoord("repeat")
                    infoOnly = rsCoord("infoOnly")
                    LD = rsCoord("LD")
                    parentID = rsCoord("parentID")
                    if status = true Then
                        completedItem_cnt = completedItem_cnt + 1
                    end if
                    if repeat = true Then
                        repeatItem_cnt = repeatItem_cnt + 1
                    end if
                    if LD = true Then
                        LDItem_cnt = LDItem_cnt + 1
                    end if
                    rsCoord.MoveNext
                LOOP
                rsCoord.Close
                SET rsCoord=nothing %>

                <% 'compare cnts to see if they match
                totalErr = false
                completeErr = false
                if coordCnt <> totalItems then
                    totalErr = true
                end if
                if completedItem_cnt <> completedItems then
                    completeErr = true
                end if
                
                if totalErr then %>
                    <h4><%=inspecCnt%>:<%=projName%>:<%=inspecDate%></h4>
                    <p>Items: [Total, Complete, Repeat, LD] [<%=coordCnt%>, <%=completedItem_cnt%>, <%=repeatItem_cnt %>, <%=LDItem_cnt %>]</p>
                    <h5 style="color: red">Error: Total Item Cnt Does Not match! [<%=totalItems %>, <%=coordCnt %>]</h5>
                    <% if fix_db then
                        updateCnt = updateCnt + 1    
                        inspectSQLUPDATE = "UPDATE Inspections SET" & _
			            " totalItems = " & coordCnt &_
			            " WHERE inspecID = " & inspecID
		                'response.Write(inspectSQLUPDATE2)
		                connSWPPP.Execute(inspectSQLUPDATE)
                    end if
                end if
                if completeErr then %>
                    <h4><%=inspecCnt%>:<%=projName%>:<%=inspecDate%></h4>
                    <p>Items: [Total, Complete, Repeat, LD] [<%=coordCnt%>, <%=completedItem_cnt%>, <%=repeatItem_cnt %>, <%=LDItem_cnt %>]</p>
                    <h5 style="color: red">Error: Completed Item Cnt Does Not match! [<%=completedItems %>, <%=completedItem_cnt %>]</h5>
                    <% if fix_db then
                        updateCnt = updateCnt + 1   
                        inspectSQLUPDATE = "UPDATE Inspections SET" & _
			            " completedItems = " & completedItem_cnt & _
			            " WHERE inspecID = " & inspecID
		                'response.Write(inspectSQLUPDATE2)
		                connSWPPP.Execute(inspectSQLUPDATE)
                    end if  
                end if 'completedItem_cnt
            end if 'rsCoord.EOF
        end if 'inspecCnt
        RS0.MoveNext
    Loop 'RSO
    RS0.Close
    SET RS0=nothing
End If 'end RSO.EOF %>
<h4>DONE - Total Reports: <%=inspecCnt%>, Updates: <%=updateCnt%></h4>
</BODY>
</HTML>
<% connSWPPP.close
SET connSWPPP=nothing %>