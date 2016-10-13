<%	response.buffer = true
	response.ContentType = "application/vnd.ms-excel"
	xFileName= QB & REPLACE(Date(),"/","_") & Request("iNum")
	response.AddHeader "content-disposition", "inline; filename="& xFileName &".xls"
err=0
xDate=Request("xDate")
IF IsDate(xDate) THEN
	xMonth=Month(xDate)
	xYear=Year(xDate)
ELSE
	err=err+1
END IF
yDate=Request("yDate")
IF NOT(IsDate(yDate)) THEN err=err+4 END IF
firstDay=CDate(xDate)
lastDay=CDate(yDate)
iNum=Request("iNum")
IF (IsNumeric(iNum) AND iNum<>"") THEN iNum=FormatNumber(iNum,0) ELSE err=err+2 END IF
IF (Request("bCycle")<5 AND Request("bCycle")>0) THEN bCycle=Request("bCycle") ELSE bCycle=1 END IF
%><!-- #include virtual="admin/connSWPPP.asp" --><%
SQL0 = "sp_qb1 '"& firstDay &"', '"& lastDay &"'," & bCycle
Set RS0 = connSWPPP.execute(SQL0) %>
<table border=0 cellpadding=0 cellspacing=0 width=2645 style='border-collapse: collapse;table-layout:fixed;width:1984pt;'>
<!--<tr><td><%= SQL0%></td></tr>-->
<tr><td>!TRNS</td><td>TRNSID</td><td>TRNSTYPE</td><td>DATE</td><td>ACCNT</td><td>NAME</td><td>CLASS</td><td>AMOUNT</td><td>DOCNUM</td><td>MEMO</td><td>CLEAR</td><td>TOPRINT</td><td>NAMEISTAXABLE</td><td>ADDR1</td><td>ADDR2</td><td>ADDR3</td><td>ADDR4</td><td>ADDR5</td><td>DUEDATE</td><td>TERMS</td><td>PAID</td><td>PAYMETH</td><td>SHIPVIA</td><td>SHIPDATE</td><td>OTHER1</td><td>REP</td><td>FOB</td><td>PONUM</td><td>INVTITLE</td><td>INVMEMO</td><td>SADDR1</td><td>SADDR2</td><td>SADDR3</td><td>SADDR4</td><td>SADDR5</td><td>PAYITEM</td><td>YEARTODATE</td><td>WAGEBASE</td><td>EXTRA</td><td>TOSEND</td><td>ISAJE</td></tr>
<tr><td>!SPL</td><td>SPLID</td><td>TRNSTYPE</td><td>DATE</td><td>ACCNT</td><td>NAME</td><td>CLASS</td><td>AMOUNT</td><td>DOCNUM</td><td>MEMO</td><td>CLEAR</td><td>QNTY</td><td>PRICE</td><td>INVITEM</td><td>PAYMETH</td><td>TAXABLE</td><td>VALADJ</td><td>REIMBEXP</td><td>SERVICEDATE</td><td>OTHER2</td><td>OTHER3</td><td>PAYITEM</td><td>YEARTODATE</td><td>WAGEBASE</td><td>EXTRA</td></tr><tr><td><nobr>!ENDTRNS</nobr></td></tr><%
currProjName = ""
DO WHILE NOT RS0.EOF
	IF RS0("projectName")<>currProjName THEN
		IF NOT(currProjName="") THEN %><tr><td>ENDTRNS</td></tr><% END IF
		currProjName=RS0("projectName")
		SQL1="SELECT SUM(inspecCost ) as sum1" &_
			" FROM vQB1 WHERE inspecDate BETWEEN '"& firstDay &"' AND '"& lastDay &"'" &_ 
			" AND projectName='"& currProjName &"'"
		SET RS1=connSWPPP.execute(SQL1)
		projectSum = RS1("sum1") 
		RS1.Close
		SET RS1=nothing 
		cnt=1 %>
<tr><td>TRNS</td><td><%= cnt %></td><td>INVOICE</td><td><%= Date() %></td><td>Accounts Receivable</td>
<td><%= TRIM(RS0("projectName"))%></td><td></td><td><%= projectSum%></td><td><%= iNum %></td><td>N</td><td>N</td><td>N</td>
<td>N</td><td><%= RS0("compName")%></td><td><%= RS0("compAddr")%></td><td><%= RS0("compAddr2")%></td><td><%= TRIM(RS0("compCity"))%><% 
IF NOT(TRIM(RS0("compCity"))="" OR TRIM(RS0("compState"))="") THEN %>,&nbsp;<% END IF %><%= TRIM(RS0("compState"))%>&nbsp;<%= TRIM(RS0("compZip"))%></td><td></td>
<td><%= Date()%></td><td></td><td>N</td><td></td><td></td><td><%= Date()%></td><td></td><td></td><td></td><td></td><td></td><td><%= TRIM(RS0("invoiceMemo")) %></td>
<td><%= RS0("projectName")%></td><td><%= TRIM(RS0("projectCity"))%><% IF NOT(TRIM(RS0("projectCity"))="" OR TRIM(RS0("projectState"))="") THEN %>,&nbsp;<% END IF %>
<%= TRIM(RS0("projectState"))%></td><td></td><td></td><td></td><td></td><td></td><td></td><td>N</td><td>N</td></tr><%
		iNum=iNum+1
	END IF 
	cnt=cnt+1 %>
<tr><td>SPL</td><td><%= cnt %></td><td>INVOICE</td><td><%= Date()%></td><td>Services</td><td></td><td></td><td>-<%= RS0("inspecCost")%></td><td></td>
<td><%= Trim(RS0("reportType")) %> Inspection <%= RS0("inspecDate")%></td><td>N</td><td></td><td><%= RS0("inspecCost")%></td><td>14</td><td></td><td>N</td><td>N</td><td>NOTHING</td></tr><%
	RS0.MoveNext
Loop %>
<tr><td>ENDTRNS</td></tr>
</table><%
RS0.Close
SET RS0=nothing
connSWPPP.Close()
SET connSWPPP=nothing 
Response.Flush()
Response.End %>