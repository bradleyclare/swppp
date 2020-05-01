<%@ language=vbscript %>
<%	response.buffer = true
	response.ContentType = "application/vnd.ms-excel"
	response.AddHeader "content-disposition", "inline; filename=QBReport.xls"
err=0
xDate=Request("xDate")
xDate=Date()
IF IsDate(xDate) THEN
	xMonth=Month(xDate)
	xYear=Year(xDate)
	firstDay=Date(xMonth &"/01/"& xYear)
	lastDay=DateAdd("d",-1,DateAdd("m",1,firstDay))
ELSE
	err=err+1
END IF
iNum=Request("iNum")
iNum=100
IF (IsNumeric(iNum) AND iNum<>"") THEN iNum=FormatNumber(iNum,0) ELSE err=err+2 END IF
'projectID=Request("pID")
'IF (IsNumeric(projectID) AND projectID<>"") THEN projectID=FormatNumber(projectID,0) ELSE err=err+4 END IF
%><!-- #include file="../admin/connSWPPP.asp" --><%
SQL0 = "SELECT inspecID, inspecDate, projectCounty, i.projectID, p.projectName" & _
	" FROM Inspections as i, Projects as p WHERE i.projectID = p.projectID" &_
	" AND i.inspecDate BETWEEN '"& firstDay &"' AND '"& lastDay &"'" &_
	" AND i.projectID IN ( SELECT projectID FROM ProjectsUsers)" &_
	" ORDER BY  p.projectName ASC, inspecDate DESC"
Set RS0 = connSWPPP.execute(SQL0) %>
<table border=0 cellpadding=0 cellspacing=0 width=2645 style='border-collapse: collapse;table-layout:fixed;width:1984pt'>
<!-- Debug info line -->
<TR><TD><%=err%></TD>
<TD><%=xDate%></TD>
<TD><%=firstDay%></TD>
<TD><%=lastDay%></TD>
<TD><%=iNum%></TD></TR>
<tr><td>!TRNS</td>
<td>TRNSID</td>
<td>TRNSTYPE</td>
<td>DATE</td>
<td>ACCNT</td>
<td>NAME</td>
<td>CLASS</td>
<td>AMOUNT</td>
<td>DOCNUM</td>
<td>MEMO</td>
<td>CLEAR</td>
<td>TOPRINT</td>
<td>NAMEISTAXABLE</td>
<td>ADDR1</td>
<td>ADDR2</td>
<td>ADDR3</td>
<td>ADDR4</td>
<td>ADDR5</td>
<td>DUEDATE</td>
<td>TERMS</td>
<td>PAID</td>
<td>PAYMETH</td>
<td>SHIPVIA</td>
<td>SHIPDATE</td>
<td>OTHER1</td>
<td>REP</td>
<td>FOB</td>
<td>PONUM</td>
<td>INVTITLE</td>
<td>INVMEMO</td>
<td>SADDR1</td>
<td>SADDR2</td>
<td>SADDR3</td>
<td>SADDR4</td>
<td>SADDR5</td>
<td>PAYITEM</td>
<td>YEARTODATE</td>
<td>WAGEBASE</td>
<td>EXTRA</td>
<td>TOSEND</td>
<td>ISAJE</td></tr>
<tr><td>!SPL</td>
<td>SPLID</td>
<td>TRNSTYPE</td>
<td>DATE</td>
<td>ACCNT</td>
<td>NAME</td>
<td>CLASS</td>
<td>AMOUNT</td>
<td>DOCNUM</td>
<td>MEMO</td>
<td>CLEAR</td>
<td>QNTY</td>
<td>PRICE</td>
<td>INVITEM</td>
<td>PAYMETH</td>
<td>TAXABLE</td>
<td>VALADJ</td>
<td>REIMBEXP</td>
<td>SERVICEDATE</td>
<td>OTHER2</td>
<td>OTHER3</td>
<td>PAYITEM</td>
<td>YEARTODATE</td>
<td>WAGEBASE</td>
<td>EXTRA</td></tr>
<tr><td>!ENDTRNS</td></tr>
<tr><td>TRNS</td>
<td>1</td>
<td>INVOICE</td>
<td>&lt;today&gt;</td>
<td>Accounts Receivable</td>
<td>&lt;company name&gt;:&lt;project name&gt;</td>
<td>&lt;total number of weeks&gt;</td>
<td>&lt;next invoice number&gt;</td>
<td>N</td>
<td>N</td>
<td>N</td>
<td>&lt;company name&gt;</td>
<td>&lt;company addr1&gt;</td>
<td>&lt;company addr2&gt;</td>
<td>&lt;company addr3&gt;</td>
<td>&lt;today&gt;</td>
<td></td>
<td>N</td>
<td></td>
<td></td>
<td>&lt;today&gt;</td>
<td></td>
<td></td>
<td></td>
<td></td>
<td></td>
<td>Due upon receipt.</td>
<td>&lt;project name&gt;</td>
<td>&lt;project city&gt;</td>
<td></td>
<td></td>
<td></td>
<td></td>
<td></td>
<td></td>
<td>N</td>
<td>N</td></tr>
<tr><td>SPL</td>
<td>2</td>
<td>INVOICE</td>
<td>&lt;today&gt;</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection <font class=font1>&lt;report date&gt;</font></td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>SPL</td>
<td>3</td>
<td>INVOICE</td>
<td>&lt;today&gt;</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection &lt;reportdate&gt;</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>SPL</td>
<td>4</td>
<td>INVOICE</td>
<td>&lt;today&gt;</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection &lt;reportdate&gt;</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>SPL</td>
<td>5</td>
<td>INVOICE</td>
<td>&lt;today&gt;</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection &lt;reportdate&gt;</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>ENDTRNS</td></tr>
<tr><td></td></tr>
<tr><td>!TRNS</td>
<td>TRNSID</td>
<td>TRNSTYPE</td>
<td>DATE</td>
<td>ACCNT</td>
<td>NAME</td>
<td>CLASS</td>
<td>AMOUNT</td>
<td>DOCNUM</td>
<td>MEMO</td>
<td>CLEAR</td>
<td>TOPRINT</td>
<td>NAMEISTAXABLE</td>
<td>ADDR1</td>
<td>ADDR2</td>
<td>ADDR3</td>
<td>ADDR4</td>
<td>ADDR5</td>
<td>DUEDATE</td>
<td>TERMS</td>
<td>PAID</td>
<td>PAYMETH</td>
<td>SHIPVIA</td>
<td>SHIPDATE</td>
<td>OTHER1</td>
<td>REP</td>
<td>FOB</td>
<td>PONUM</td>
<td>INVTITLE</td>
<td>INVMEMO</td>
<td>SADDR1</td>
<td>SADDR2</td>
<td>SADDR3</td>
<td>SADDR4</td>
<td>SADDR5</td>
<td>PAYITEM</td>
<td>YEARTODATE</td>
<td>WAGEBASE</td>
<td>EXTRA</td>
<td>TOSEND</td>
<td>ISAJE</td></tr>
<tr><td>!SPL</td>
<td>SPLID</td>
<td>TRNSTYPE</td>
<td>DATE</td>
<td>ACCNT</td>
<td>NAME</td>
<td>CLASS</td>
<td>AMOUNT</td>
<td>DOCNUM</td>
<td>MEMO</td>
<td>CLEAR</td>
<td>QNTY</td>
<td>PRICE</td>
<td>INVITEM</td>
<td>PAYMETH</td>
<td>TAXABLE</td>
<td>VALADJ</td>
<td>REIMBEXP</td>
<td>SERVICEDATE</td>
<td>OTHER2</td>
<td>OTHER3</td>
<td>PAYITEM</td>
<td>YEARTODATE</td>
<td>WAGEBASE</td>
<td>EXTRA</td></tr>
<tr><td>!ENDTRNS</td></tr>
<tr><td>TRNS</td>
<td>1</td>
<td>INVOICE</td>
<td>1/26/2004</td>
<td>Accounts Receivable</td>
<td>Ryland Homes:R.H. of Texas L.P.:Ames Meadow</td>
<td>400</td>
<td>320</td>
<td></td>
<td>N</td>
<td>N</td>
<td>N</td>
<td>R.H. of Texas L.P.</td>
<td>17855 North Dallas Pkwy</td>
<td>Suite 200</td>
<td>Dallas, TX 75287</td>
<td>1/26/2004</td>
<td></td>
<td>N</td>
<td></td>
<td></td>
<td>1/26/2004</td>
<td></td>
<td></td>
<td></td>
<td></td>
<td></td>
<td>Due upon receipt.</td>
<td>Ames Meadow Phase IIIA</td>
<td>Lancaster, TX</td>
<td></td>
<td></td>
<td></td>
<td></td>
<td></td>
<td></td>
<td>N</td>
<td>N</td></tr>
<tr><td>SPL</td>
<td>2</td>
<td>INVOICE</td>
<td>1/26/2004</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection 01/02/04</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>SPL</td>
<td>3</td>
<td>INVOICE</td>
<td>1/26/2004</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection 01/08/04</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>SPL</td>
<td>4</td>
<td>INVOICE</td>
<td>1/26/2004</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection 01/15/04</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>SPL</td>
<td>5</td>
<td>INVOICE</td>
<td>1/26/2004</td>
<td>Services</td>
<td></td>
<td></td>
<td>-100</td>
<td></td>
<td>Weekly Inspection 01/22/04</td>
<td>N</td>
<td></td>
<td>100</td>
<td>14</td>
<td></td>
<td>N</td>
<td>N</td>
<td>NOTHING</td></tr>
<tr><td>ENDTRNS</td></tr>
</table></div>