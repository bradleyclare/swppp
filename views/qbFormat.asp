<%
'	response.buffer = true
'	response.ContentType = "application/vnd.ms-excel"
'	response.AddHeader "content-disposition", "inline; filename=dynamicTest.xls"
err=0
xDate=Request("xDate")
IF IsDate(xDate) THEN
	xMonth=Month(xDate)
	xYear=Year(xDate)
	firstDay=CDate(xMonth &"/01/"& xYear)
	lastDay=DateAdd("d",-1,DateAdd("m",1,firstDay))
ELSE
	err=err+1
END IF
iNum=Request("iNum")
IF (IsNumeric(iNum) AND iNum<>"") THEN iNum=FormatNumber(iNum,0) ELSE err=err+2 END IF
%><!-- #include file="../admin/connSWPPP.asp" --><%
SQL0 = "SELECT inspecID, inspecDate, projectCounty, i.projectID, p.projectName," &_
	" i.compName, compAddr, compAddr2, projectCity" &_
	" FROM Inspections as i, Projects as p WHERE i.projectID = p.projectID" &_
	" AND i.inspecDate BETWEEN '"& firstDay &"' AND '"& lastDay &"'" &_
	" ORDER BY  p.projectName ASC, inspecDate DESC"
Response.Write(SQL0)
Set RS0 = connSWPPP.execute(SQL0) %>
<table border=1 cellpadding=0 cellspacing=2 width=2645 style='border-collapse: collapse;table-layout:fixed;width:1984pt;'>
<tr><td><%= err %></td><td><%= xDate %></td><td><%= firstDay %></td><td><%= lastDay %></td><td><%= iNum %></td></tr><%
currProjID = ""
DO WHILE NOT RS0.EOF
	IF RS0("projectID")<>currProjID THEN
		IF NOT(currProjID="") THEN %><tr><td>ENDTRNS</td></tr><tr><td></td></tr><% END IF
		currProjID=RS0("projectID")
		SQL1="SELECT COUNT(*) FROM Inspections WHERE inspecDate BETWEEN '"& firstDay &"' AND '"& lastDay &"' AND projectID="& currProjID
		SET RS1=connSWPPP.execute(SQL1)
		totNumWeeks = RS1(0) 
		SET RS1=nothing 
		cnt=1 %>
<tr><td>!TRNS</td><td>TRNSID</td><td>TRNSTYPE</td><td>DATE</td><td>ACCNT</td><td>NAME</td><td>CLASS</td><td>AMOUNT</td><td>DOCNUM</td><td>MEMO</td><td>CLEAR</td><td>TOPRINT</td><td>NAMEISTAXABLE</td><td>ADDR1</td><td>ADDR2</td><td>ADDR3</td><td>ADDR4</td><td>ADDR5</td><td>DUEDATE</td><td>TERMS</td><td>PAID</td><td>PAYMETH</td><td>SHIPVIA</td><td>SHIPDATE</td><td>OTHER1</td><td>REP</td><td>FOB</td><td>PONUM</td><td>INVTITLE</td><td>INVMEMO</td><td>SADDR1</td><td>SADDR2</td><td>SADDR3</td><td>SADDR4</td><td>SADDR5</td><td>PAYITEM</td><td>YEARTODATE</td><td>WAGEBASE</td><td>EXTRA</td><td>TOSEND</td><td>ISAJE</td></tr>
<tr><td>!SPL</td><td>SPLID</td><td>TRNSTYPE</td><td>DATE</td><td>ACCNT</td><td>NAME</td><td>CLASS</td><td>AMOUNT</td><td>DOCNUM</td><td>MEMO</td><td>CLEAR</td><td>QNTY</td><td>PRICE</td><td>INVITEM</td><td>PAYMETH</td><td>TAXABLE</td><td>VALADJ</td><td>REIMBEXP</td><td>SERVICEDATE</td><td>OTHER2</td><td>OTHER3</td><td>PAYITEM</td><td>YEARTODATE</td><td>WAGEBASE</td><td>EXTRA</td></tr><tr><td><nobr>!ENDTRNS</nobr></td></tr>
<tr><td>TRNS</td><td><%= cnt %></td><td>INVOICE</td><td><%= Date() %></td><td>Accounts Receivable</td>
<td><%= RS0("compName")%>:<%= RS0("projectName")%></td><td></td><td><%= (totNumWeeks * 100)%></td><td><%= iNum %></td><td>N</td><td>N</td><td>N</td>
<td><%= RS0("compName")%></td><td><%= RS0("compAddr")%></td><td><%= RS0("compAddr2")%></td><td><% '--compAddr3 does not exist%></td>
<td><%= Date()%></td><td></td><td>N</td><td></td><td></td><td><%= Date()%></td><td></td><td></td><td></td><td></td><td></td><td>Thank you for your business.</td>
<td><%= RS0("projectName")%></td><td><%= RS0("projectCity")%></td><td></td><td></td><td></td><td></td><td></td><td></td><td>N</td><td>N</td></tr><%
		iNum=iNum+1
	END IF 
	cnt=cnt+1 %>
<tr><td>SPL</td><td><%= cnt %></td><td>INVOICE</td><td><%= Date()%></td><td>Services</td><td></td><td></td><td>-100</td><td></td>
<td>Weekly Inspection <%= RS0("inspecDate")%></td><td>N</td><td></td><td>100</td><td>14</td><td></td><td>N</td><td>N</td><td>NOTHING</td></tr><%
	RS0.MoveNext
Loop %>
<tr><td>ENDTRNS</td></tr>
</table><%
SET RS0=nothing
connSWPPP.Close()
SET connSWPPP=nothing %>