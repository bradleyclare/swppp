<%Response.Buffer = False%>
<%
If Not Session("validAdmin") And Not Session("validInspector") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
inspecID = Session("inspecID")
IF Request("inspecID")<>"" THEN 
	inspecID = Request("inspecID") 
	Session("inspecID")=inspecID
END IF
%><!-- #include file="../connSWPPP.asp" --><%
If Request.Form.Count > 0 Then	
	'Response.Write("Form Submitted")
	Function strQuoteReplace(strValue)
		strQuoteReplace = Replace(strValue, "'", "''")
	End Function	
	if Request.Form("submit_optional_btn") = "documents" then
		Response.Redirect("editReportOptionalLinks.asp?inspecID=" + inspecID)
	End If
	userID = Session("userID")
	inspector = TRIM(strQuoteReplace(Request("inspector")))	
	If inspector <> "" Then userID = inspector End If
	vscr = TRIM(strQuoteReplace(Request("vscr")))
	if vscr = "" Then vscr = 0
	ldscr = TRIM(strQuoteReplace(Request("ldscr")))
	if ldscr = "" Then ldscr = 0	
	bmps=Request("bmpsInPlace")
	sediment=Request("sediment")
	upProjPhase= strQuoteReplace(Request("projectPhase"))
	
	inspecDate = strQuoteReplace(Request("inspecDate"))
	If inspecDate = "1/1/1900" Then
	   Response.Write("Fixing bad date. Line 35.\n")
		inspecDate = date()
	End If
	inspectSQLUPDATE = "UPDATE Inspections SET" & _
		" inspecDate = '" & inspecDate & "'" & _
		", projectName = '" & strQuoteReplace(Request("projectName")) & "'" 
		IF LEN(TRIM(upProjPhase))=0 THEN 
			inspectSQLUPDATE=inspectSQLUPDATE &", projectPhase=null" 
		ELSE
			inspectSQLUPDATE=inspectSQLUPDATE &", projectPhase='" & upProjPhase &"'"
		END IF
		includeItems = 0
		if Request("includeItems") = "on" then includeItems = 1 End If
	   openItemAlert = 0
		if Request("openItemAlert") = "on" then openItemAlert = 1 End If	
      systemic = 0
		if Request("systemic") = "on" then systemic = 1 End If	
      compliance = 0
		if Request("compliance") = "on" then compliance = 1 End If
		hortonQuestions = 0
		if Request("horton") = "on" then hortonQuestions = 1 End If
		hortonSignV = 0
		if Request("hortonV") = "on" then hortonSignV = 1 End If
		hortonSignLD = 0
		if Request("hortonLD") = "on" then hortonSignLD = 1 End If
		forestarQuestions = 0
		if Request("forestar") = "on" then forestarQuestions = 1 End If
		inspectSQLUPDATE=inspectSQLUPDATE &", projectAddr = '" & strQuoteReplace(Request("projectAddr")) & "'" & _
		", projectCity = '" & strQuoteReplace(Request("projectCity")) & "'" & _
		", projectState = '" & Request("projectState") & "'" & _
		", projectZip = '" & strQuoteReplace(Request("projectZip")) & "'" & _
		", projectCounty = '" & Request("projectCounty") & "'" & _
		", onsiteContact = '" & strQuoteReplace(Request("onsiteContact")) & "'" & _
		", officePhone = '" & strQuoteReplace(Request("officePhone")) & "'" & _
		", emergencyPhone = '" & strQuoteReplace(Request("emergencyPhone")) & "'" & _
		", reportType = '" & Request("reportType") & "'" & _
		", inches = " & Request("inches") & _
		", bmpsInPlace = " & bmps & _
		", sediment = " & sediment & _
		", userID = " & inspector & _
		", compName = '" & strQuoteReplace(Request("compName")) & "'" & _
		", compAddr = '" & strQuoteReplace(Request("compAddr")) & "'" & _
		", compAddr2 = '" & strQuoteReplace(Request("compAddr2")) & "'" & _
		", compCity = '" & strQuoteReplace(Request("compCity")) & "'" & _
		", compState = '" & Request("compState") & "'" & _
		", compZip = '" & strQuoteReplace(Request("compZip")) & "'" & _
		", compPhone = '" & strQuoteReplace(Request("compPhone")) & "'" & _
		", compContact = '" & strQuoteReplace(Request("compContact")) & "'" & _
		", contactPhone = '" & strQuoteReplace(Request("contactPhone")) & "'" & _
		", contactFax = '" & strQuoteReplace(Request("contactFax")) & "'" & _
		", contactEmail = '" & strQuoteReplace(Request("contactEmail")) & "'" & _
		", includeItems = " & includeItems & _
        ", openItemAlert = " & openItemAlert & _
		", compliance = " & compliance & _
        ", systemic = " & systemic & _
        ", systemicNote = '" & strQuoteReplace(Request("systemicNote")) & "'" & _
		", horton = " & hortonQuestions & _
		", hortonSignV = " & hortonSignV & _
		", hortonSignLD = " & hortonSignLD & _
		", vscr = " & vscr & _
		", ldscr = " & ldscr & _
		", forestar = " & forestarQuestions & _
		" WHERE inspecID = " & inspecID
        'response.Write(inspectSQLUPDATE)
	    connSWPPP.Execute(inspectSQLUPDATE)
    
		totalItems = 0
		completedItems = 0
        inspecDate = strQuoteReplace(Request("inspecDate"))
		If inspecDate = "1/1/1900" Then
		   Response.Write("Fixing bad date. Line 102.\n")
			inspecDate = date()
		End If
		for n = 1 to 999 step 1
            'Response.Write("coord:coID:" & CStr(n)&":"& Request("coord:coID:" & CStr(n)) &"<br/>")
		    if Trim(Request("coord:coID:" & CStr(n))) = "" then
		        exit for
		    end if
            '-- dbo.spAEDCoordinate @_iCOID int, @_DelFlag smallint, @_inspecID int, @_iCoordinates char(50), @_icorrectiveMods char(255), @_iOrderBy int
            DelCoord = 0
            if Request("coord:del:"& CStr(n)) = "on" then 
				DelCoord = 1 
			ElseIf IsNumeric(Request("coord:orderby:"& CStr(n))) then
				totalItems = totalItems + 1
                If compliance Then
                    message = "Site is in Compliance checked with actions defined! Either uncheck Site is in Compliance or remove the actions and resubmit."
                    Response.Write("<script language=""javascript"">alert(""" + message + """);</script>")
                End If
			End If
			Complete = 0
			if Request("coord:status:"& CStr(n)) = "on" then 
				Complete = 1 
				completedItems = completedItems + 1
			End If
			Repeat = 0
			if Request("coord:repeat:"& CStr(n)) = "on" then 
                Repeat = 1
                parentID = Request("coord:parentID:"& CStr(n)) 'keep the parentID the same if repeat item
            Else
                parentID = Request("coord:coID:"& CStr(n)) 'set parentID to current coID, this will be zero for a new item and the current coID for a carryover item
            End If
			useAddress = 0
            if Request("coord:useAddress:"& Cstr(n)) = "on" then useAddress = 1 End If
			address = TRIM(strQuoteReplace(Request("coord:addressName:"& Cstr(n))))
			locationName = TRIM(strQuoteReplace(Request("coord:locationName:"& Cstr(n))))
            infoOnly = 0
			if Request("coord:infoOnly:"& CStr(n)) = "on" then 
                totalItems = totalItems - 1
                if Complete Then
                    completedItems = completedItems - 1
                End If
                infoOnly = 1 
        	End If
        	LD = 0
        	if Request("coord:LD:"& CStr(n)) = "on" then LD = 1 End If
        	NLN = 0
        	if Request("coord:NLN:"& CStr(n)) = "on" then NLN = 1 End If
			pond = 0
        	if Request("coord:pond:"& CStr(n)) = "on" then pond = 1 End If
			sedloss = 0
        	if Request("coord:sedloss:"& CStr(n)) = "on" then sedloss = 1 End If
			sedlossw = 0
        	if Request("coord:sedlossw:"& CStr(n)) = "on" then sedlossw = 1 End If
			ce = 0
        	if Request("coord:ce:"& CStr(n)) = "on" then ce = 1 End If
			street = 0
        	if Request("coord:street:"& CStr(n)) = "on" then street = 1 End If
			sfeb = 0
        	if Request("coord:sfeb:"& CStr(n)) = "on" then sfeb = 1 End If
			rockdam = 0
        	if Request("coord:rockdam:"& CStr(n)) = "on" then rockdam = 1 End If
			ip = 0
        	if Request("coord:ip:"& CStr(n)) = "on" then ip = 1 End If
			wo = 0
        	if Request("coord:wo:"& CStr(n)) = "on" then wo = 1 End If
			veg = 0
        	if Request("coord:veg:"& CStr(n)) = "on" then veg = 1 End If
			stock = 0
        	if Request("coord:stock:"& CStr(n)) = "on" then stock = 1 End If
			toilet = 0
        	if Request("coord:toilet:"& CStr(n)) = "on" then toilet = 1 End If
			trash = 0
        	if Request("coord:trash:"& CStr(n)) = "on" then trash = 1 End If
			dewater = 0
        	if Request("coord:dewater:"& CStr(n)) = "on" then dewater = 1 End If
			discharge = 0
        	if Request("coord:discharge:"& CStr(n)) = "on" then discharge = 1 End If
			dust = 0
        	if Request("coord:dust:"& CStr(n)) = "on" then dust = 1 End If
			riprap = 0
        	if Request("coord:riprap:"& CStr(n)) = "on" then riprap = 1 End If
			outfall = 0
        	if Request("coord:outfall:"& CStr(n)) = "on" then outfall = 1 End If
			intop = 0
        	if Request("coord:intop:"& CStr(n)) = "on" then intop = 1 End If
			swalk = 0
        	if Request("coord:swalk:"& CStr(n)) = "on" then swalk = 1 End If
			mormix = 0
        	if Request("coord:mormix:"& CStr(n)) = "on" then mormix = 1 End If
			ada = 0
        	if Request("coord:ada:"& CStr(n)) = "on" then ada = 1 End If
			dway = 0
        	if Request("coord:dway:"& CStr(n)) = "on" then dway = 1 End If
			flume = 0
        	if Request("coord:flume:"& CStr(n)) = "on" then flume = 1 End If
			osc = 0
			if Request("coord:osc:"& CStr(n)) = "on" then osc = 1 End If
			AssignDate = inspecDate
        	if Repeat = 1 Then
			   	AssignDate = Request("coord:assignDate:"& CStr(n))
			End If
			OrderBy = 0
        	if IsNumeric(Request("coord:orderby:"& CStr(n))) then OrderBy = Request("coord:orderby:"& CStr(n)) End If
			'SQLc = SQLc &"/*<br/>*/Exec spAEDCoordinate "& Request("coord:coID:"& CStr(n)) &", "& DelCoord &", "& inspecID &", '"& Replace(Request("coord:coord:"& CStr(n)),"--","�") &"', '"& Replace(Request("coord:mods:"& CStr(n)),"--","�") &"', "& OrderBy &";"
			SQLc = SQLc &"/*<br/>*/Exec spAEDCoordinate "& _ 
			Request("coord:coID:"& CStr(n)) &", "& _ 
			DelCoord &", "& _ 
			inspecID &", '"& _ 
			Replace(Request("coord:coord:"& CStr(n)),"--","�") &"', '"& _ 
			Replace(Request("coord:mods:"& CStr(n)),"--","�") &"', "& _ 
			OrderBy &", '"& _ 
			AssignDate &"', '"& _ 
			Request("coord:completeDate:"& CStr(n)) &"', "& _ 
			Complete &", " & _ 
			Repeat &", " & _ 
			useAddress &", '" & _ 
			address &"', '" & _
			locationName &"', " & _
        	infoOnly &", " & _
        	LD &", " & _
        	NLN &", " & _
        	parentID &", " & _
        	pond &", " & _
			sedloss &", " & _
			sedlossw &", " & _
			ce &", " & _
			street &", " & _
			sfeb &", " & _
			rockdam &", " & _
			ip &", " & _
			wo &", " & _
			veg &", " & _
			stock &", " & _
			toilet &", " & _
			trash &", " & _
			dewater &", " & _
			dust &", " & _
			riprap &", " & _
			outfall &", " & _
			intop &", " & _
			swalk &", " & _
			mormix &", " & _
			ada &", " & _
			dway &", " & _
			flume &", " & _
			osc &", " & _ 
			discharge &";"
		next
		'Response.Write(SQLc & "</br>")	
        if Len(SQLc) > 0 then connSWPPP.execute(SQLc) end if
'Response.End	

	'update items counts
	inspectSQLUPDATE2 = "UPDATE Inspections SET" & _
		" totalItems = " & totalItems & _
		", completedItems = " & completedItems & _
		" WHERE inspecID = " & inspecID
'response.Write(inspectSQLUPDATE2)
	connSWPPP.Execute(inspectSQLUPDATE2)

    if Request.Form("submit_view_reports_btn") = "view reports" then
		connSWPPP.Close
	    Set connSWPPP = Nothing
    	Response.Redirect("viewReports.asp")
	End If
	If request("submit") = "Edit Report & Project Info" Then	
	    connSWPPP.Close
	    Set connSWPPP = Nothing
    	Response.Redirect("viewReports.asp")
    End If
End If
	inspecSQLSELECT = "SELECT inspecDate, i.projectName, i.projectPhase, projectAddr, projectCity, projectState" & _
		", projectZip, projectCounty, onsiteContact, officePhone, emergencyPhone, i.projectID, compName, released" & _
		", compAddr, compAddr2, compCity, compState, compZip, compPhone, compContact, contactPhone, contactFax" & _
		", contactEmail, reportType, inches, bmpsInPlace, sediment, userID, includeItems, compliance, totalItems" & _
		", completedItems, openItemAlert, systemic, systemicNote, horton, hortonSignV, hortonSignLD, vscr, ldscr" & _
		", forestar, p.collectionName" & _
		" FROM Inspections as i, Projects as p" & _
		" WHERE i.projectID = p.projectID AND inspecID = " & inspecID
'--Response.Write(inspecSQLSELECT & "<br>")
	Set rsReport = connSWPPP.execute(inspecSQLSELECT)
'baseDir = "d:\vol\swpppinspections.com\www\htdocs\" 
baseDir = "D:\Inetpub\wwwroot\SWPPP\"%>
<html>
<head>
	<title>SWPPP INSPECTIONS : edit inspection report</title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
	<STYLE>
	    select.long {
	        font-size: xx-small;
	    }
	</STYLE>
	<script type="text/javascript" language="JavaScript" src="../js/validReports.js"></script>
	<script type="text/javascript" language="JavaScript" src="../js/validReports1.2.js"></script>
	<link href="../../css/jquery-ui.min.css" rel="stylesheet" type="text/css"/>
	<link href="../../css/jquery-ui.structure.min.css" rel="stylesheet" type="text/css"/>
	<link href="../../css/jquery-ui.theme.min.css" rel="stylesheet" type="text/css"/>
	<script src="../../js/jquery.js" type="text/javascript"></script>
	<script src="../../js/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript" >
    $(function () {
        $(".datepicker").datepicker();
    });
</script>
<script type="text/javascript" >
    $(document).ready(function () {
        $('#dialog-confirm').dialog({
            autoOpen: false,
            resizable: false,
            height: "auto",
            width: 500,
            modal: true,
            buttons: {
                "Delete All Items": function () {
                    //check all delete checkboxes (coord:deleteX)
                    var i;
                    for (i = 1; i < 99; ++i) {
                        var e = document.getElementsByName("coord:del:" + i);
                        if (e.length) {
                            $("[name='coord:del:" + i + "']")[0].checked = true;
                        } else {
                            break;
                        }
                    }
                    $('#compliance-checkbox')[0].checked = true;
                    $(this).dialog("close");
                    document.getElementById("theForm").submit();
                },
                Cancel: function () {
                    $(this).dialog("close");
                }
            }
        });

        $('#compliance-checkbox').click(
             function () {
                if ($('#compliance-checkbox').is(":checked")) {
                   $("#dialog-confirm").dialog('open');
                   return false;
                } else {
                   $('#compliance-checkbox')[0].checked = false;
                   document.getElementById("theForm").submit();
                }
             }
         );

        $('#systemic-confirm').dialog({
           autoOpen: false,
           resizable: false,
           height: "auto",
           width: 500,
           modal: true,
           buttons: {
              "Save Note": function () {
                 //submit form to save note and checkbox
                 var s_ta = document.getElementsByName("systemic_note");
                 var s_h = document.getElementsByName("systemicNote");
                 s_h[0].value = s_ta[0].value;
                 $('#systemic-checkbox')[0].checked = true;
                 $(this).dialog("close");
                 document.getElementById("theForm").submit();
              },
              Cancel: function () {
                 $(this).dialog("close");
              }
           }
        });

        $('#systemic-checkbox').click(
             function () {
                if ($('#systemic-checkbox').is(":checked")) {
                   $("#systemic-confirm").dialog('open');
                   return false;
                } else {
                   $('#systemic-checkbox')[0].checked = false;
                   document.getElementById("theForm").submit();
                }
             }
          );

        $('#includeItems-checkbox').click(
            function () {
                document.getElementById("theForm").submit();
            }
        );

        $('#openItemAlert-checkbox').click(
            function () {
                document.getElementById("theForm").submit();
            }
        );
    });
</script>
	<script type="text/javascript" language="JavaScript1.2"><!--
    // we Can't just use the same transfer function for both directions because
    // the hidden input keys off of the t2 value solely...-->

    function addOption(t1, t2, t3) {
        var index = t3.selectedIndex;
        if (index > -1) {
            var newoption = new Option(t3.options[index].text, t3.options[index].value, true, true);
            t2.options[t2.length] = newoption;
            if (!document.getElementById) history.go(0);
            t3.options[index] = null;
            t3.selectedIndex = 0;
            var tempStr = "";
            for (var i = 0; i < (t2.length) ; i++) {
                tempStr = tempStr + (t2.options[i].value) + ":";
            }
            t1.value = tempStr;
        }
    } function delOption(t1, t3, t2) {
        var index = t3.selectedIndex;
        if (index > -1) {
            var newoption = new Option(t3.options[index].text, t3.options[index].value, true, true);
            t2.options[t2.length] = newoption;
            if (!document.getElementById) history.go(0);
            t3.options[index] = null;
            t3.selectedIndex = 0;
            var tempStr = "";
            for (var i = 0; i < (t3.length) ; i++) {
                tempStr = tempStr + (t3.options[i].value) + ":";
            }
            t1.value = tempStr;
        }
    }
    function swapOption(t1, t2, slideDir) {
        var curIndex = t2.selectedIndex;
        var swapIndex = curIndex;
        var maxIndex = t2.length;
        if (curIndex > -1) {
            (slideDir == "up") ? (swapIndex = curIndex - 1) : (swapIndex = curIndex + 1);
            if ((swapIndex > -1) && (swapIndex < t2.length)) {
                var newOption = new Option(t2.options[swapIndex].text, t2.options[swapIndex].value, true, true);
                t2.options[maxIndex] = newOption;
                t2.options[swapIndex].text = t2.options[curIndex].text;
                t2.options[swapIndex].value = t2.options[curIndex].value;
                t2.options[curIndex].text = t2.options[maxIndex].text;
                t2.options[curIndex].value = t2.options[maxIndex].value;
                t2.options[maxIndex] = null;
                t2.selectedIndex = swapIndex;
                var tempStr = "";
                for (var i = 0; i < (t2.length) ; i++) {
                    tempStr = tempStr + (t2.options[i].value) + ":";
                }
                t1.value = tempStr;
            }
        }
    }
    function editNarrative(inspID) {
        var basePath = "http://www.swppp.com";
        var URL = "/admin/maintain/editNarrative.asp?inspecID=" + inspID;
        var params = "height=420,width=520,status=yes,toolbar=no,menubar=no, directories=no, location=no, scrollbars=no, resizable=no";
        window.open(URL, "", params);
    }

	function hortonQuestions(inspID) {
        var basePath = "http://www.swppp.com";
        var URL = "/admin/maintain/hortonQuestions.asp?inspecID=" + inspID;
        var params = "";
        window.open(URL, "", params);
    }

	 function forestarQuestions(inspID) {
        var basePath = "http://www.swppp.com";
        var URL = "/admin/maintain/forestarQuestions.asp?inspecID=" + inspID;
        var params = "";
        window.open(URL, "", params);
    }

    function useAddressLookup(obj) {
        var parts = obj.name.split(":");
        var selectname = "coord:locationName:" + parts[2];
        var s = document.getElementsByName(selectname);
        var selectname2 = "coord:addressName:" + parts[2];
        var s2 = document.getElementsByName(selectname2);
        var selectname3 = "coord:coord:" + parts[2];
        var s3 = document.getElementsByName(selectname3);
        if (obj.checked) //enable select object
        {
            s[0].className = "";
            s2[0].className = "";
            s3[0].className = "hide";
        }
        else //disable select object
        {
            s[0].className = "hide";
            s2[0].className = "hide";
            s3[0].className = "";
        }
    }

    function setSelectValue(obj) {
        //selected value of addOptions dropdown
        var val = obj.selectedIndex;
        var parts = obj.name.split(":");

        //find address dropdown list to set the same value
        var selectname = "coord:address:" + parts[2];
        var s = document.getElementsByName(selectname);
        s[0].selectedIndex = val;

        //set the hidden object to keep address name
        var hiddenname2 = "coord:addressName:" + parts[2];
        var s2 = document.getElementsByName(hiddenname2);
        s2[0].value = s[0].value.trim();

        //set the hidden object for locationName
        var hiddenname3 = "coord:locationName:" + parts[2];
        var s3 = document.getElementsByName(hiddenname3);
        s3[0].value = obj.value.trim();
    }

    function displayAddressSelect(obj) {
        var parts = obj.name.split(":");
        var num = parts[2];

        var pos = getPosition(obj);

        //display the select div
        var s1 = document.getElementsByName("addressOptionsPopup");
        s1[0].className = "addressOptionsPopup show";
        s1[0].style.top = pos.y;
        s1[0].style.left = pos.x;

        //set the hidden div in the select div to remember what number we are modifying
        var s2 = document.getElementsByName("currentAddressNum");
        s2[0].value = num;
    }

    function setAddress(obj) {

        //get number from hidden div
        var s1 = document.getElementsByName("currentAddressNum");
        var num = s1[0].value;

        //get the dropdown options
        var sl = document.getElementsByName("locationOptions");
        var selectedName = sl[0].value;

        //get address dropdown options
        var sa = document.getElementsByName("addressOptions");
        sa[0].selectedIndex = sl[0].selectedIndex;
        var selectedAddress = sa[0].value;

        //set the hidden object to keep address name
        var hiddenname2 = "coord:addressName:" + num;
        var s3 = document.getElementsByName(hiddenname2);
        s3[0].value = selectedAddress;

        //set the hidden object for locationName
        var hiddenname3 = "coord:locationName:" + num;
        var s4 = document.getElementsByName(hiddenname3);
        s4[0].value = selectedName;

        //hide the select div
        var s0 = document.getElementsByName("addressOptionsPopup");
        s0[0].className = "addressOptionsPopup hide";
    }

    function close_popup() {
        //hide the select div
        var s0 = document.getElementsByName("addressOptionsPopup");
        s0[0].className = "addressOptionsPopup hide";
    }

    function getPosition(el) {
        var xPos = 0;
        var yPos = 0;
 
        while (el) {
            if (el.tagName == "BODY") {
                // deal with browser quirks with body/window/document and page scroll
                var xScroll = el.scrollLeft || document.documentElement.scrollLeft;
                var yScroll = el.scrollTop || document.documentElement.scrollTop;
 
                xPos += (el.offsetLeft - xScroll + el.clientLeft);
                yPos += (el.offsetTop - yScroll + el.clientTop);
            } else {
                // for all other non-BODY elements
                xPos += (el.offsetLeft - el.scrollLeft + el.clientLeft);
                yPos += (el.offsetTop - el.scrollTop + el.clientTop);
            }
 
            el = el.offsetParent;
        }
        return {
            x: xPos,
            y: yPos
        };
    }

    function displayCommonItemSelect(obj) {
        var parts = obj.name.split(":");
        var num = parts[2];

        var pos = getPosition(obj);

        //display the select div
        var s1 = document.getElementsByName("commonItemsPopup");
        s1[0].className = "commonItemsPopup show";
        s1[0].style.top  = pos.y;
        s1[0].style.left = pos.x;

        //set the hidden div in the select div to remember what number we are modifying
        var s2 = document.getElementsByName("commonItemsNum");
        s2[0].value = num;
    }

    function setCommonItem(obj) {

        //get number from hidden div
        var s1 = document.getElementsByName("commonItemsNum");
        var num = s1[0].value;

        //get the dropdown options
        var sl = document.getElementsByName("commonItemOptions");
        var selectedItem = sl[0].value;

        //set the hidden object to keep address name
        var hiddenname2 = "coord:mods:" + num;
        var s3 = document.getElementsByName(hiddenname2);
        s3[0].value = selectedItem;

        //hide the select div
        var s0 = document.getElementsByName("commonItemsPopup");
        s0[0].className = "commonItemsPopup hide";

        sl[0].selectedIndex = 0;
    }

    function close_item_popup() {
        //hide the select div
        var s0 = document.getElementsByName("commonItemsPopup");
        s0[0].className = "commonItemsPopup hide";
    }

	function horton_checkbox(state) {
		//display confirm dialog
      var chbx_id = document.getElementById("horton-checkbox");
		if (state) {
			var box= confirm("Are you sure you want activate DR Horton questions from this report?");
			if (box==true) {
				//change state and submit form
				chbx_id.checked = true;
				document.getElementById("theForm").submit();
			} else {
				chbx_id.checked = false;
			}
		} else {
			var box= confirm("Are you sure you want to remove DR Horton questions for this report?");
			if (box==true) {
				//change state and submit form
				chbx_id.checked = false;
				document.getElementById("theForm").submit();
			} else {
				chbx_id.checked = true;
			}
		}
	 }

	 function forestar_checkbox(state) {
		//display confirm dialog
      var chbx_id = document.getElementById("forestar-checkbox");
		if (state) {
			var box= confirm("Are you sure you want activate Forestar questions from this report?");
			if (box==true) {
				//change state and submit form
				chbx_id.checked = true;
				document.getElementById("theForm").submit();
			} else {
				chbx_id.checked = false;
			}
		} else {
			var box= confirm("Are you sure you want to remove Forestar questions for this report?");
			if (box==true) {
				//change state and submit form
				chbx_id.checked = false;
				document.getElementById("theForm").submit();
			} else {
				chbx_id.checked = true;
			}
		}
	 }

	 function checkbox(id, state) {
		//display confirm dialog
      var chbx_id = id;
		if (state) {
			//change state and submit form
			chbx_id.checked = true;
			document.getElementById("theForm").submit();
		} else {
			//change state and submit form
			chbx_id.checked = false;
			document.getElementById("theForm").submit();
		}
	 }

</script>
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>edit report</h1>
<form id="theForm" method="post" action="<%=Request.ServerVariables("script_name")%>?inspecID=<%=inspecID%>" onsubmit="return isReady(this)";>
	<input type="hidden" name="inspecID" value="<%=inspecID %>" />
	<input type="hidden" name="projectID" value="<%=rsReport("projectID") %>" />
   <input type="hidden" name="systemicNote" value="<%=Trim(rsReport("systemicNote"))%>" />
	
<div id="dialog-confirm" title="Confirm Compliance">
	<p>Site in Compliance? What do you want to do with the open items?</p>
</div>
<div id="systemic-confirm" title="Systemic Note">
	<p>Site has systemic issues. Please enter note detailing issue.</p>
   <textarea name="systemic_note" cols="65" rows="3"><%=Trim(rsReport("systemicNote"))%></textarea>
</div>
<table width="90%">
<tr><th width="30%" align="center">report date</th><th width="30%" align="center">project name</th><th width="30%" align="center">customer name</th></tr>

<% inspecDate = Trim(rsReport("inspecDate"))
If inspecDate = "1/1/1900" Then
	Response.Write("Fixing bad date. Line 685.\n")
	inspecDate = date()
End If 
allowUpdate = True
if Session("validInspector") and not Session("validAdmin") and rsReport("released") Then
	allowUpdate = False
End If
%>

<tr><td align="center"><% = inspecDate %></td><td align="center"><% = Trim(rsReport("projectName")) %>&nbsp<% = Trim(rsReport("projectPhase")) %></td><td align="center"><% = Trim(rsReport("compName")) %></td></tr>
<tr>
<% If allowUpdate Then %>
<td align="center"><a href="deleteReport.asp?inspecID=<%=inspecID %>"><button type="button">delete report</button></a></td>
<td align="center"><a href="releasereports_test.asp?inspecID=<%=inspecID%>&projID=<%=rsReport("projectID")%>"><button type="button">view email report</button></a></td>
<% If Session("validAdmin") then %>
<td align="center"><a href="manage_addresses.asp?inspecID=<%=inspecID%>"><button type="button">upload address information</button></a></td>
<% End If %>
</tr>
<% End If %>
</table>
<br/>
<hr>
<% If allowUpdate Then %>
<table width="90%" border="0" cellpadding="2" cellspacing="0">
	<tr><td align="right" bgcolor="#eeeeee"><b>type of report:</b></td>
			<td bgcolor="#999999"><select name="reportType">
<% 	SQL2="SELECT * FROM ReportTypes where priority > 0 ORDER BY priority DESC, reportTypeID ASC"
	SET RS2=connSWPPP.execute(SQL2)
	DO WHILE NOT RS2.EOF %><option value="<%= Trim(RS2("reportType"))%>"<% 
	If Trim(rsReport("reportType")) = TRIM(RS2("reportType")) Then %> selected<% End If %>><%= Trim(RS2("reportType"))%></option>
<% 	RS2.MoveNext
	LOOP %>	</select></td>
		</tr>
	<%	'admin can change inspector name
If Session("validAdmin") Then
	insSQLSELECT = "SELECT DISTINCT u.userID, firstName, lastName" & _
		" FROM Users as u, ProjectsUsers as pu" & _
		" WHERE u.userID = pu.userID AND (pu.rights='inspector' OR pu.rights='admin')" &_
		" ORDER BY lastName"
	Set connUser = connSWPPP.execute(insSQLSELECT) %>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>inspector:</strong></td>
			<td bgcolor="#999999"><select name="inspector">
				<% Do While Not connUser.EOF %><option value="<%= connUser("userID") %>" <% If rsReport("userID")=connUser("userID") Then %>selected<% End If %>><%= Trim(connUser("firstName")) & "&nbsp;" & Trim(connUser("lastName")) %></option> <%= rsReport("userID") %>*
<%					connUser.moveNext
				Loop				
	connUser.Close
	Set connUser = Nothing %>
			</select></td></tr>
<%	Else 
	SQLa="SELECT * FROM Users WHERE userID="& rsReport("userID") 
	Set connUser= connSWPPP.execute(SQLa) %>
		<tr> 
			<td align="right" bgcolor="#eeeeee"><strong>inspector:</strong></td>
			<td bgcolor="#999999"><%= Trim(connUser("firstName"))%> <%=Trim(connUser("lastName"))%>
				<INPUT type="hidden" name="inspector" value="<%= rsReport("userID")%>"></td></tr>
<%	End If %>
	<tr><td align="right" bgcolor="#eeeeee"><b>apply DR Horton questions:</b></td>
	<td bgcolor="#999999">Yes?
	<% If rsReport("horton") = True Then %>
		<input id="horton-checkbox" type="checkbox" name="horton" onChange="horton_checkbox(0)" checked/> 
		<% SQLA = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
		Set RSA = connSWPPP.execute(SQLA)
    	If RSA.EOF Then %>
			<input type="button" value="view questions" style="background-color: #f44336;" onClick="hortonQuestions('<%= inspecID%>')">
		<% Else %>
			<input type="button" value="view questions" onClick="hortonQuestions('<%= inspecID%>')">
		<% End If %>
		V Sign?
		<% If rsReport("hortonSignV") = True Then %>
		   <input id="horton-checkbox-sign-V" type="checkbox" name="hortonV" onChange="checkbox(this.id, 0)" checked/>
			<select name="vscr">
				<% 'show listbox of available approvers 
				rightsSELECT = "SELECT DISTINCT u.userID, firstName, lastName" & _
					" FROM Users as u, ProjectsUsers as pu" & _
					" WHERE u.userID = pu.userID AND pu.projectID = " & rsReport("projectID") & _
					" AND pu.rights='vscr' ORDER BY lastName"
				Set connRights = connSWPPP.execute(rightsSELECT) %>
				<option value="0">None</option>
				<% Do While Not connRights.EOF %>
					<option value="<%= connRights("userID") %>" <% If rsReport("vscr")=connRights("userID") Then %>selected<% End If %>><%= Trim(connRights("firstName")) & "&nbsp;" & Trim(connRights("lastName")) %></option>
				<%	connRights.moveNext
				Loop %>
			</select>
		<% Else %>
		   <input id="horton-checkbox-sign-V" type="checkbox" name="hortonV" onChange="checkbox(this.id, 1)" />
	   <% End If %>
		LD Sign?
		<% If rsReport("hortonSignLD") = True Then %>
		   <input id="horton-checkbox-sign-LD" type="checkbox" name="hortonLD" onChange="checkbox(this.id, 0)" checked/>
			<select name="ldscr">
				<% 'show listbox of available approvers 
				rightsSELECT = "SELECT DISTINCT u.userID, firstName, lastName" & _
					" FROM Users as u, ProjectsUsers as pu" & _
					" WHERE u.userID = pu.userID AND pu.projectID = " & rsReport("projectID") & _
					" AND pu.rights='ldscr' ORDER BY lastName"
				Set connRights = connSWPPP.execute(rightsSELECT) %>
				<option value="0">None</option>
				<% Do While Not connRights.EOF %>
					<option value="<%= connRights("userID") %>" <% If rsReport("ldscr")=connRights("userID") Then %>selected<% End If %>><%= Trim(connRights("firstName")) & "&nbsp;" & Trim(connRights("lastName")) %></option>
				<%	connRights.moveNext
				Loop %>
			</select>
		<% Else %>
		   <input id="horton-checkbox-sign-LD" type="checkbox" name="hortonLD" onChange="checkbox(this.id, 1)" />
	   <% End If %>
	<% Else %>
		<input id="horton-checkbox" type="checkbox" name="horton" onChange="horton_checkbox(1)" />
	<% End If %>
	</td></tr>

	<tr><td align="right" bgcolor="#eeeeee"><b>apply Forestar questions:</b></td>
	<td bgcolor="#999999">Yes?
	<% If rsReport("forestar") = True Then %>
		<input id="forestar-checkbox" type="checkbox" name="forestar" onChange="forestar_checkbox(0)" checked/> 
		<% SQLA = "SELECT * FROM HortonAnswers WHERE inspecID = " & inspecID
		Set RSA = connSWPPP.execute(SQLA)
    	If RSA.EOF Then %>
			<input type="button" value="view questions" style="background-color: #f44336;" onClick="forestarQuestions('<%= inspecID%>')">
		<% Else %>
			<input type="button" value="view questions" onClick="forestarQuestions('<%= inspecID%>')">
		<% End If %>
		SSCR Sign
		<select name="ldscr">
			<% 'show listbox of available approvers 
			rightsSELECT = "SELECT DISTINCT u.userID, firstName, lastName" & _
				" FROM Users as u, ProjectsUsers as pu" & _
				" WHERE u.userID = pu.userID AND pu.projectID = " & rsReport("projectID") & _
				" AND pu.rights='ldscr' ORDER BY lastName"
			Set connRights = connSWPPP.execute(rightsSELECT) %>
			<option value="0">None</option>
			<% Do While Not connRights.EOF %>
				<option value="<%= connRights("userID") %>" <% If rsReport("ldscr")=connRights("userID") Then %>selected<% End If %>><%= Trim(connRights("firstName")) & "&nbsp;" & Trim(connRights("lastName")) %></option>
			<%	connRights.moveNext
			Loop %>
		</select>
	<% Else %>
		<input id="forestar-checkbox" type="checkbox" name="forestar" onChange="forestar_checkbox(1)" />
	<% End If %>
	</td></tr>
</table>
<% End If %>

<!------------------------------------- Coordinates --------------------------- --->
<hr/>
<% totalItems = rsReport("totalItems")
completedItems = rsReport("completedItems")
if totalItems <> "" and totalItems <> 0 Then
	score = FormatNumber((completedItems/totalItems)*100,0) & "%" 
Else
	score = "N/A"
End If%>
<table width="100%">
<tr>
<% If allowUpdate Then %>
<td colspan="2" align="center"><input name="submit_coord_btn_top" type="submit" style="font-size: 20px;" value="submit"/></td>
<% End If %>
<td colspan="3" align="center"><h2>action items</h2></td>
</tr>
<% If allowUpdate Then %>
<tr width="13%"><td>total items: <%=totalItems%></td><td width="13%">completed items: <%=completedItems%></td><td width="13%">report score:<%=score%></td><td width="13%">site is in compliance
<% If rsReport("compliance") = True Then %>
	<input id="compliance-checkbox" type="checkbox" name="compliance" checked/>
<% Else %>
	<input id="compliance-checkbox" type="checkbox" name="compliance" />
<% End If %>
</td><td width="13%" align="left">apply scoring to report
<% If rsReport("includeItems") = True Then %>
	<input id='includeItems-checkbox' type="checkbox" name="includeItems" checked/>
<% Else %>
	<input id='includeItems-checkbox' type="checkbox" name="includeItems" />
<% End If %>
</td><td width="13%" align="left">open item alert
<% If rsReport("openItemAlert") = True Then %>
	<input id='openItemAlert-checkbox' type="checkbox" name="openItemAlert" checked/>
<% Else %>
	<input id='openItemAlert-checkbox' type="checkbox" name="openItemAlert" />
<% End If %>
</td><td width="13%" align="left">alert
<% If rsReport("systemic") = True Then %>
	<input id='systemic-checkbox' type="checkbox" name="systemic" checked/>
<% Else %>
	<input id='systemic-checkbox' type="checkbox" name="systemic" />
<% End If %>
</td></tr>
<% End If %>
</table><br/>
<% coordSQLSELECT = "SELECT * FROM Coordinates WHERE inspecID=" & inspecID & " ORDER BY orderby"	
'Response.Write(coordSQLSELECT)
Set rsCoord = connSWPPP.execute(coordSQLSELECT)
addressSQLSELECT = "SELECT addressID, locationName, address FROM Addresses WHERE projectID=" & rsReport("projectID") & " ORDER BY locationName"
'Response.Write(addressSQLSELECT)
Set rsAddress = connSWPPP.execute(addressSQLSELECT)
'create single popup list to display when user wants to modify the address
locationName1 = ""
addressName1 = ""
%>
<div class="addressOptionsPopup hide" name="addressOptionsPopup">
<h3>Select coordinates here:</h3>
<input type="hidden" name="currentAddressNum" value="1" />
<input type="hidden" name="commonItemsNum" value="1" />
<select name="locationOptions" onchange="setAddress(this)">
<% if not rsAddress.EOF Then
    cnt = 0
	Do While Not rsAddress.EOF 
        cnt = cnt + 1
        if (cnt = 1) Then
            locationName1 = TRIM(rsAddress("locationName")) 
        End If
		name = TRIM(rsAddress("locationName")) %>
		<option value="<%=name%>"><%=name%></option>
	<% rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst
End If %>
</select>
<select name="addressOptions" class="hide" readonly >
<% if not rsAddress.EOF Then
    cnt = 0
	Do While Not rsAddress.EOF 
        cnt = cnt + 1
        if (cnt = 1) Then
			addressName1 = TRIM(rsAddress("address"))
        End If
		name = TRIM(rsAddress("address")) %>
		<option value="<%=name%>"><%=name%></option>
	<% rsAddress.MoveNext
	Loop 
	rsAddress.MoveFirst
End If %>
</select>
<br /><br />
<input type="button" onclick="close_popup()" value="Close Window" />
</div>
<% 
itemSQLSELECT = "SELECT itemID, itemName FROM CommonItems ORDER BY itemName"
'Response.Write(addressSQLSELECT)
Set rsItems = connSWPPP.execute(itemSQLSELECT) 
%>
<div class="commonItemsPopup hide" name="commonItemsPopup">
<h3>select common item:</h3>
<input type="hidden" name="commonItemNum" value="1" />
<select name="commonItemOptions" onchange="setCommonItem(this)">
<option value=""></option>
<% if not rsItems.EOF Then
    cnt = 0
	Do While Not rsItems.EOF 
        cnt = cnt + 1
		item = TRIM(rsItems("itemName")) %>
		<option value="<%=item%>"><%=item%></option>
	<% rsItems.MoveNext
	Loop 
	rsItems.MoveFirst
End If %>
</select>
<br /><br />
<input type="button" onclick="close_item_popup()" value="Close Window" />
</div>
<center>
    Click "Repeat" on all items that you want the assign date to stay the same. All other items will be updated to the current date on SUBMIT.
    <table><tr>
    <td><a href="../../views/openActionItems.asp?pID=<%=rsReport("projectID")%>" target="_blank">open items page</a></td>
    <td><a href="../../views/completedActionItems.asp?pID=<%=rsReport("projectID")%>" target="_blank">completed items page</a></td>
    </tr></table>
</center><br/>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
<% 
'If rsCoord.EOF Then
'	Response.Write("<tr><td colspan='2' align='center'><i>There is no data at this time.</i></td></tr>")		
'Else
    n = 1
	Do While Not rsCoord.EOF
	    coID = rsCoord("coID")
		correctiveMods = Trim(rsCoord("correctiveMods"))
		orderby = rsCoord("orderby")
		coordinates = Trim(rsCoord("coordinates"))
		existingBMP = Trim(rsCoord("existingBMP")) 
		assignDate = rsCoord("assignDate") 
		completeDate = rsCoord("completeDate")
		status = rsCoord("status")
		repeat = rsCoord("repeat")
		useAddress = rsCoord("useAddress")
		address = TRIM(rsCoord("address"))
		locationName = TRIM(rsCoord("locationName"))
      infoOnly = rsCoord("infoOnly")
      LD = rsCoord("LD")
      NLN = rsCoord("NLN")
      parentID = rsCoord("parentID")
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
		osc = rsCoord("osc")
		discharge = rsCoord("discharge")
        if isNull(parentID) or parentID = "" then 'initialize the parentID if never set
            parentID = coID
        end if
		'Response.Write("ID: " & coID & ", Coord: " & coordinates & ", LocName: " & locationName & ", address: " & address & ", Mods: " & correctiveMods & "<br/>") 
		%>
	<input type="hidden" name="coord:coID:<%= n %>" value="<%= coID %>" />
	<!--<input type="hidden" name="coord:status:<%= n %>" value="<%= status %>" />-->
	<input type="hidden" name="coord:completeDate:<%= n %>" value="" />
   <input type="hidden" name="coord:NLN:<%= n %>" value="<%=NLN %>"/>
   <input type="hidden" name="coord:parentID:<%= n %>" value="<%=parentID %>" />
	<tr><td>id#</td>
	<td><%= coID %></td>
	<td>address<input type="checkbox" name="coord:useAddress:<%= n %>" onclick="useAddressLookup(this)" 
	<% if (useAddress) = True Then %>
		 checked
	<% End If %>
	/></td>
	<td colspan="3">
    <input type="text" size="40" name="coord:locationName:<%= n %>" onclick="displayAddressSelect(this)" value="<%=locationName %>"
	<% if (useAddress) = False Then %>
		class="hide"
	<% End If %>
	/></td>
	<td colspan="3">
    <input type="text" size="40" name="coord:addressName:<%= n %>" value="<%=address%>"
	<% if (useAddress) = False Then %>
		class="hide"
	<% End If %>
	readonly /></td>    
    </tr>
	<tr><td>order</td>
	<td><input type="text" name="coord:orderby:<%= n %>" size="10" value="<% = orderby %>" /></td>
	<td>location info:</td>
	<td colspan="6"><input name="coord:coord:<%= n %>" type="text" value="<%= coordinates %>" size="100%"  
	<% if (useAddress) = True Then %>
		class="hide"
	<% End If %>
	></td></tr>
	<tr><td>delete<input type="checkbox" name="coord:del:<%= n %>" /></td>
	<td>Repeat
	<% If repeat = True Then %>
		<input type="checkbox" name="coord:repeat:<%= n %>" checked/>
	<% Else %>
		<input type="checkbox" name="coord:repeat:<%= n %>" />
	<% End If %>
	</td><td>note
	<% If infoOnly = True Then %>
		<input type="checkbox" name="coord:infoOnly:<%= n %>" checked/>
	<% Else %>
		<input type="checkbox" name="coord:infoOnly:<%= n %>" />
	<% End If %>
	</td>
	<td rowspan="3" colspan="6"><textarea name="coord:mods:<%= n %>" cols="100%" rows="5"><%= correctiveMods %></textarea></td></tr>
	<tr><td>assignDate</td>
	<td><input <% If Session("validAdmin") Then %> class=datepicker <% End If %> type="text" name="coord:assignDate:<%= n %>" size="10" value="<%= assignDate %>" <% If not Session("validAdmin") Then %> readonly <% End If %>/></td>
	<td><input type="button" onclick="displayCommonItemSelect(this)" name="coord:item:<%=n%>" value="common item" /></td></tr>
    <tr><td>OSC
	 <% If osc = True Then %>
			<input type="checkbox" name="coord:osc:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:osc:<%=n %>" />
		<% End If %>
    </td><td>
	 LD
	<% If LD = True Then %>
		<input type="checkbox" name="coord:LD:<%= n %>" checked/>
	<% Else %>
		<input type="checkbox" name="coord:LD:<%= n %>" />
	<% End If %>
   </td>
    <td> status
    <% If status = True Then %>
        <input type="checkbox" name="coord:status:<%=n %>" checked />
    <% Else %>
        <input type="checkbox" name="coord:status:<%=n %>" />
    <% End If %>
    </td></tr>
<%	IF existingBMP <> "-1" THEN %>
	<tr>
		<td align="right"><b>Existing BMP:</b></td>
		<td><font face="Times" size="2.5pt"><%= existingBMP %></font></td>
	</tr>
<% 	END IF %>

<%	IF rsReport("horton") = True or rsReport("forestar") = True THEN %>
	<tr>
		<td colspan="9">
		<table>
		<td> ada
		<% If ada = True Then %>
			<input type="checkbox" name="coord:ada:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:ada:<%=n %>" />
		<% End If %>
		</td>
		<% If rsReport("horton") = True Then %>
		<td> ce
		<% If ce = True Then %>
			<input type="checkbox" name="coord:ce:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:ce:<%=n %>" />
		<% End If %>
		</td>
		<% End If %>
		<td> dewater
		<% If dewater = True Then %>
			<input type="checkbox" name="coord:dewater:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:dewater:<%=n %>" />
		<% End If %>
		</td>
		<td> dis
		<% If discharge = True Then %>
			<input type="checkbox" name="coord:discharge:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:discharge:<%=n %>" />
		<% End If %>
		</td>
		<td> dust
		<% If dust = True Then %>
			<input type="checkbox" name="coord:dust:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:dust:<%=n %>" />
		<% End If %>
		</td>
		<% If rsReport("horton") = True Then %>
		<td> dway
		<% If dway = True Then %>
			<input type="checkbox" name="coord:dway:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:dway:<%=n %>" />
		<% End If %>
		</td>
		<% End If %>
		<td> eb/ft/sf
		<% If sfeb = True Then %>
			<input type="checkbox" name="coord:sfeb:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:sfeb:<%=n %>" />
		<% End If %>
		</td>
		<% If rsReport("horton") = True Then %>
		<td> flume
		<% If flume = True Then %>
			<input type="checkbox" name="coord:flume:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:flume:<%=n %>" />
		<% End If %>
		</td>
		<% End If %>
		<td> intop
		<% If intop = True Then %>
			<input type="checkbox" name="coord:intop:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:intop:<%=n %>" />
		<% End If %>
		</td>
		<td> ip
		<% If ip = True Then %>
			<input type="checkbox" name="coord:ip:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:ip:<%=n %>" />
		<% End If %>
		</td>
		<td> mormix
		<% If mormix = True Then %>
			<input type="checkbox" name="coord:mormix:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:mormix:<%=n %>" />
		<% End If %>
		</td>
		<% If rsReport("horton") = True Then %>
		<td> outfall
		<% If outfall = True Then %>
			<input type="checkbox" name="coord:outfall:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:outfall:<%=n %>" />
		<% End If %>
		</td>
		<% End If %>
		<td> pond
		<% If pond = True Then %>
			<input type="checkbox" name="coord:pond:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:pond:<%=n %>" />
		<% End If %>
		</td>
		<% If rsReport("horton") = True Then %>
		<td> riprap
		<% If riprap = True Then %>
			<input type="checkbox" name="coord:riprap:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:riprap:<%=n %>" />
		<% End If %>
		</td>
		<% End If %>
		<td> rockdam
		<% If rockdam = True Then %>
			<input type="checkbox" name="coord:rockdam:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:rockdam:<%=n %>" />
		<% End If %>
		</td>
		<td> sedloss
		<% If sedloss = True Then %>
			<input type="checkbox" name="coord:sedloss:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:sedloss:<%=n %>" />
		<% End If %>
		</td>
		<td> sedlossw
		<% If sedlossw = True Then %>
			<input type="checkbox" name="coord:sedlossw:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:sedlossw:<%=n %>" />
		<% End If %>
		</td>
		<% If rsReport("horton") = True Then %>
		<td> stock
		<% If stock = True Then %>
			<input type="checkbox" name="coord:stock:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:stock:<%=n %>" />
		<% End If %>
		</td>
		<% End If %>
		<td> street
		<% If street = True Then %>
			<input type="checkbox" name="coord:street:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:street:<%=n %>" />
		<% End If %>
		</td>
		<td> swalk
		<% If swalk = True Then %>
			<input type="checkbox" name="coord:swalk:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:swalk:<%=n %>" />
		<% End If %>
		</td>
		<td> toilet
		<% If toilet = True Then %>
			<input type="checkbox" name="coord:toilet:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:toilet:<%=n %>" />
		<% End If %>
		</td>
		<td> trash
		<% If trash = True Then %>
			<input type="checkbox" name="coord:trash:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:trash:<%=n %>" />
		<% End If %>
		</td>
		<td> veg
		<% If veg = True Then %>
			<input type="checkbox" name="coord:veg:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:veg:<%=n %>" />
		<% End If %>
		</td>
		<td> wo
		<% If wo = True Then %>
			<input type="checkbox" name="coord:wo:<%=n %>" checked />
		<% Else %>
			<input type="checkbox" name="coord:wo:<%=n %>" />
		<% End If %>
		</td>
		</table>
		</td>
	</tr>
<% 	END IF %>
	<tr><td colspan="9"><hr align="center" width="100%" size="1"></td></tr>
<%	 	rsCoord.MoveNext
        n = n + 1
	Loop 	
'End If ' END No Results Found
rsCoord.Close
Set rsCoord = Nothing %>
<% for m = n to n+9 step 1 %>
	<input type="hidden" name="coord:coID:<%= m %>" value="0" />
	<input type="hidden" name="coord:del:<%= m %>" value="0" />
	<input type="hidden" name="coord:completeDate:<%= m %>" value="" />
	<!--<input type="hidden" name="coord:status:<%= m %>" value="0" />-->
	<input type="hidden" name="coord:repeat:<%= m %>" value="0" />
   <input type="hidden" name="coord:NLN:<%= m %>" value="0" />
   <input type="hidden" name="coord:parentID:<%= m %>" value="0" />
	<tr><td>id#</td>
	<td>0</td>
	<td colspan="3">address<input type="checkbox" name="coord:useAddress:<%= m %>" onclick="useAddressLookup(this)"/></td>
	<td>
   <input type="text" name="coord:locationName:<%= m %>" onclick="displayAddressSelect(this)" value="<%=locationName1 %>" class="hide" /></td>
	<td colspan="3">
	<%temp = addressName1 %>
   <input type="text" name="coord:addressName:<%= m %>" value="<%=temp%>" class="hide" readonly /></td></tr>
	<tr><td>order</td>
	<td><input type="text" name="coord:orderby:<%= m %>" size="10" value="" /></td>
	<td>location:</td>
	<td colspan="6"><input name="coord:coord:<%= m %>" type="text" value="" size="100%" ></td></tr>
	<tr><td></td>
	<td></td>
	<td>note <input type="checkbox" name="coord:infoOnly:<%= m %>" /></td>
	<td rowspan="3" colspan="6"><textarea name="coord:mods:<%= m %>" cols="100%" rows="5"></textarea></td></tr>
	<tr><td>assignDate</td>
	<td><input <% If Session("validAdmin") Then %> class=datepicker <% End If %> type="text" name="coord:assignDate:<%= m %>" size="10" value="" readonly /></td>
	<td><input type="button" onclick="displayCommonItemSelect(this)" name="coord:item:<%=m%>" value="common item" /></td></tr>
	<tr><td>OSC <input type="checkbox" name="coord:osc:<%= m %>" /> </td>
	<td>LD <input type="checkbox" name="coord:LD:<%= m %>" /></td></tr>
	<%	IF rsReport("horton") = True or rsReport("forestar") = True THEN %>
	<tr>
		<td colspan="9">
		<table>
		<td> ada <input type="checkbox" name="coord:ada:<%=m %>" /></td>
		<% If rsReport("horton") = True Then %>
		<td> ce <input type="checkbox" name="coord:ce:<%=m %>" /></td>
		<% End If %>
		<td> dewater <input type="checkbox" name="coord:dewater:<%=m %>" /></td>
		<td> dis <input type="checkbox" name="coord:dis:<%=m %>" /></td>
		<td> dust <input type="checkbox" name="coord:dust:<%=m %>" /></td>
		<% If rsReport("horton") = True Then %>
		<td> dway <input type="checkbox" name="coord:dway:<%=m %>" /></td>
		<% End If %>
		<td> eb/ft/sf <input type="checkbox" name="coord:sfeb:<%=m %>" /></td>
		<% If rsReport("horton") = True Then %>
		<td> flume <input type="checkbox" name="coord:flume:<%=m %>" /></td>
		<% End If %>
		<td> intop <input type="checkbox" name="coord:intop:<%=m %>" /></td>
		<td> ip <input type="checkbox" name="coord:ip:<%=m %>" /></td>
		<td> mormix <input type="checkbox" name="coord:mormix:<%=m %>" /></td>
		<% If rsReport("horton") = True Then %>
		<td> outfall <input type="checkbox" name="coord:outfall:<%=m %>" /></td>
		<% End If %>
		<td> pond <input type="checkbox" name="coord:pond:<%=m %>" /></td>
		<% If rsReport("horton") = True Then %>
		<td> riprap <input type="checkbox" name="coord:riprap:<%=m %>" /></td>
		<% End If %>
		<td> rockdam <input type="checkbox" name="coord:rockdam:<%=m %>" /></td>
		<td> sedloss <input type="checkbox" name="coord:sedloss:<%=m %>" /></td>
		<td> sedlossw <input type="checkbox" name="coord:sedlossw:<%=m %>" /></td>
		<% If rsReport("horton") = True Then %>
		<td> stock <input type="checkbox" name="coord:stock:<%=m %>" /></td>
		<% End If %>
		<td> street <input type="checkbox" name="coord:street:<%=m %>" /></td>
		<td> swalk <input type="checkbox" name="coord:swalk:<%=m %>" /></td>
		<td> toilet <input type="checkbox" name="coord:toilet:<%=m %>" /></td>
		<td> trash <input type="checkbox" name="coord:trash:<%=m %>" /></td>
		<td> veg <input type="checkbox" name="coord:veg:<%=m %>" /></td>
		<td> wo <input type="checkbox" name="coord:wo:<%=m %>" /></td>
		</table>
		</td>
	</tr>
<% 	END IF %>
	<tr><td colspan="9"><hr align="center" width="100%" size="1"></td></tr>
<% next %>
	<% If allowUpdate Then %>
	<tr><td colspan="5" align="center"><br>
	<input name="submit_coord_btn" type="submit" style="font-size: 20px;" value="submit"/>
	<br><br></td></tr>
	<% End If %>
</table>
<% rsAddress.Close 
Set rsAddress = Nothing %>

<hr>
<h2>Project Information: #<%=rsReport("projectID")%></h2>
<table width="90%" border="0" cellpadding="2" cellspacing="0">
		<!-- date -->
		<tr><td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Date:</b></td>
			<td bgcolor="#999999"> <input type="text" name="inspecDate" size="10" value="<% = inspecDate %>" <% If not Session("validAdmin") Then %> readonly <% End If %>> <small>&nbsp;(mm / dd / yyyy)</small></td>
		</tr>
		<!-- project name -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Project Name | Phase:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectName" size="50" value="<% = Trim(rsReport("projectName")) %>" <% If not Session("validAdmin") Then %> readonly <% End If %>/>
			<input type="text" name="projectPhase" size="20" value="<% = Trim(rsReport("projectPhase")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>/></td>
		</tr>
        <tr><td align="right" bgcolor="#eeeeee"><b>Project Group:</b></td>
			<td bgcolor="#999999"><%=Trim(rsReport("collectionName"))%></td>
		</tr>
		<!-- project location -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Project Location:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectAddr" size="50" value="<% = Trim(rsReport("projectAddr")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>> </td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>City, State, Zip:</b></td>
			<td bgcolor="#999999"><input type="text" name="projectCity" size="20" value="<% = Trim(rsReport("projectCity")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>> &nbsp; 
      
		<%If not Session("validAdmin") Then%>
		<input type="hidden" name="projectState" value="<% =Trim(rsReport("projectState"))%>">
		<%End If%>
		<select name="projectState" <%If not Session("validAdmin") Then%> disabled <%End If%>>
		<% SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
		SET RS0=connSWPPP.execute(SQL0)
		IF IsNull(TRIM(rsReport("projectState"))) THEN rsReport("projectState")="TX" END IF
		DO WHILE NOT RS0.EOF %>	
			<option value="<%= RS0("stateAbbr")%>"
			<% If Trim(rsReport("projectState")) = RS0("stateAbbr") Then %> selected<% End If %>><%= Trim(RS0("stateAbbr"))%></option>
			<%	RS0.MoveNext
		LOOP %>	
		</select> &nbsp;

		<input type="text" name="projectZip" size="5" value="<% = Trim(rsReport("projectZip")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>> </td>
		</tr>
		<!-- onsite contact -->
		<tr><td align="right" bgcolor="#eeeeee"><b>County:</b></td>
		<td bgcolor="#999999">

		<%If not Session("validAdmin") Then%>
		<input type="hidden" name="projectCounty" value="<% =Trim(rsReport("projectCounty"))%>">
		<%End If%>	
		<select name="projectCounty" <%If not Session("validAdmin") Then%> disabled <%End If%>>
      <option value=""></option>
		<% SQL1="SELECT * FROM Counties ORDER BY priority DESC, countyName ASC"
		SET RS1=connSWPPP.execute(SQL1)
		DO WHILE NOT RS1.EOF %>
			<option value="<%= Trim(RS1("countyName"))%>"<% 
			If Trim(rsReport("projectCounty")) = TRIM(RS1("countyName")) Then %> selected<% End If %>><%= Trim(RS1("countyName"))%></option>
			<%	RS1.MoveNext
		LOOP %>	
		</select></td>
		
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="onsiteContact" size="50" value="<% = Trim(rsReport("onsiteContact")) %>"></td>
		</tr>
		<!-- office # -->
		<tr><td align="right" bgcolor="#eeeeee"><b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input name="officePhone" type="text" size="50" value="<% = Trim(rsReport("officePhone")) %>"></td>
		</tr>
		<!-- emergency # -->
		<tr><td align="right" bgcolor="#eeeeee"> <b>On-Site Contact:</b></td>
			<td bgcolor="#999999"><input name="emergencyPhone" type="text" size="50" value="<% = Trim(rsReport("emergencyPhone")) %>"></td>
		</tr><tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr>
</table>

<h2>Company Information</h2>
<table width="90%" border="0" cellpadding="2" cellspacing="0">
		<tr><td width="35%" bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td width="55%" bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Company Name:</b></td>
			<td bgcolor="#999999"><input type="text" name="compName" size="50" value="<% = Trim(rsReport("compName")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr>
		<!-- Address -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Address 1:</b></td>
			<td bgcolor="#999999"><input name="compAddr" type="text" size="50" value="<% = Trim(rsReport("compAddr")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Address 2:</b></td>
			<td bgcolor="#999999"><input name="compAddr2" type="text" size="50" value="<% = Trim(rsReport("compAddr2")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>City:</b></td>
			<td bgcolor="#999999"><input name="compCity" type="text" size="20" value="<% = Trim(rsReport("compCity")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>State:</b></td>
			<td bgcolor="#999999">

		<%If not Session("validAdmin") Then%>
		<input type="hidden" name="compState" value="<% =Trim(rsReport("compState"))%>">
		<%End If%>		
		<select name="compState" <%If not Session("validAdmin") Then%> disabled <%End If%>>
		<% SQL0="SELECT * FROM States ORDER BY priority DESC, stateName ASC"
		SET RS0=connSWPPP.execute(SQL0)
		IF IsNull(TRIM(rsReport("compState"))) THEN rsReport("compState")="TX" END IF
		DO WHILE NOT RS0.EOF %>
			<option value="<%= Trim(RS0("stateAbbr"))%>"<% 
			If Trim(rsReport("compState")) = RS0("stateAbbr") Then %> selected<% End If %>><%= Trim(RS0("stateAbbr"))%></option>
			<%	RS0.MoveNext
		LOOP %>	
		</select></td>
		
		</tr><tr><td align="right" bgcolor="#eeeeee"><b>Zip:</b></td>
			<td bgcolor="#999999"><input name="compZip" type="text" size="5" value="<% = Trim(rsReport("compZip")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr>
		<!-- main telephone number -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Company Phone:</b></td>
			<td bgcolor="#999999"><input name="compPhone" type="text" size="20" value="<% = Trim(rsReport("compPhone")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr>
		<!-- contact -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact:</b></td>
			<td bgcolor="#999999"><input type="text" name="compContact" size="50" value="<% = Trim(rsReport("compContact")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr>
		<!-- phone -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact Phone:</b></td>
			<td bgcolor="#999999"><input name="contactPhone" type="text" size="20" value="<% = Trim(rsReport("contactPhone")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr>
		<!-- fax -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact Fax:</b></td>
			<td bgcolor="#999999"><input name="contactFax" type="text" size="20" value="<% = Trim(rsReport("contactFax")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr>
		<!-- e-mail -->
		<tr><td align="right" bgcolor="#eeeeee"><b>Contact E-Mail:</b></td>
			<td bgcolor="#999999"><input type="text" name="contactEmail" size="50" value="<% = Trim(rsReport("contactEmail")) %>" <%If not Session("validAdmin") Then%> readonly <%End If%>></td>
		</tr><tr><td bgcolor="#eeeeee"><img src="../../images/dot.gif" width="5" height="5"></td>
			<td bgcolor="#999999"><img src="../../images/dot.gif" width="5" height="5"></td>
		</tr><tr><td colspan="2"><img src="../../images/dot.gif" width="5" height="5"></td>
		<!-- Type of Report? Weekly, Storm, Complaint, ? -->
		<!-- Rain -->
<% IF Trim(rsReport("inches")) > "-1" THEN %>
		<tr><td align="right" bgcolor="#eeeeee"><b>Inches of Rain:</b></td>
			<td bgcolor="#999999"><input type="text" name="inches" size="6"	value="<% = Trim(rsReport("inches")) %>"></td></tr>
<% ELSE %>
		<INPUT type="hidden" name="inches" value="<%= Trim(rsReport("inches"))%>">
<% END IF %>
		<!-- BMPs? y/n -->
<% IF rsReport("bmpsInPlace") = "-1" THEN %>
		<INPUT type="hidden" name="bmpsInPlace" value="<%= rsReport("bmpsInPlace")%>">
<% ELSE %>
		<tr><td align="right" bgcolor="#eeeeee"><b>Are BMPs in place?</b></td>
			<td bgcolor="#999999"> <select name="bmpsInPlace">
					<option value="1"<% If rsReport("bmpsInPlace")="1" Then %> selected<% End If %>>Yes</option>
					<option value="0"<% If rsReport("bmpsInPlace")="0" Then %> selected<% End If %>>No</option>
				</select></td></tr>
<% END IF %>
		<!-- sediment loss or pollution? y/n -->
<% IF rsReport("bmpsInPlace") = "-1" THEN %>
		<INPUT type="hidden" name="sediment" value="<%= rsReport("sediment")%>">
<% ELSE %>
		<tr><td align="right" bgcolor="#eeeeee"><b>Sediment Loss or Pollution?</b></td>
			<td bgcolor="#999999"><select name="sediment">
					<option value="1"<% If rsReport("sediment")="1" Then %> selected<% End If %>>Yes</option>
					<option value="0"<% If rsReport("sediment")="0" Then %> selected<% End If %>>No</option>
				</select></td></tr>
<% END IF %>
</Table>

<!-- ------------- Optional Links ----------------------------------------------------- -->

<% If Session("validAdmin") Then %>
<hr/>
<center><input name="submit_optional_btn" type="submit" style="font-size: 20px;" value="documents"/></center>
<% End If %>
<% If allowUpdate Then %>
<hr/>
<center><input name="submit_view_reports_btn" type="submit" style="font-size: 20px;" value="view reports"/></center>
<% End If %>

<!------------------------------------- Images ---------------------------------------->

<% IF False and NOT(Session("noImages")) THEN %>
	<hr/>
	<h2>Images</h2>
<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0"><%
smImgSQLSELECT = "SELECT imageID, smallImage, description FROM Images WHERE inspecID=" & inspecID	
Set rsSmImages = connSWPPP.execute(smImgSQLSELECT)

If rsSmImages.EOF Then
	Response.Write("<tr><td colspan='2' align='center'><i>There are " & _
		"no images at this time.</i></td></tr>")
Else %> 
	<tr><td colspan="3">Edit an image or description by selecting the name.<br><br></td></tr>
	<tr>
	<% Do While Not rsSmImages.EOF
	imageID = rsSmImages("imageID")
	smallImage = Trim(rsSmImages("smallImage"))
	desc = Trim(rsSmImages("description"))
	
	iDataRows = iDataRows + 1
	If iDataRows > 3 Then
		Response.Write("	</tr>" & VBCrLf & "	<tr>")
		iDataRows = 1
	End If %>
	<td height="125"><div align="center"><a href="editImage.asp?imageID=<%= imageID %>&inspecID=<%=inspecID%>"><%= desc %></a><br>
	<a href="editImage.asp?imageID=<%= imageID %>">
	<img src="../../images/sm/<%= smallImage %>" alt="<%= smallImage %>" width="100" height="75" 
		border="0"></a></div></td>
	<% rsSmImages.MoveNext
	Loop %>
	</tr>
	<% End If

	rsSmImages.Close 
	Set rsSmImages = Nothing %>
	<tr><td colspan="3" align="center"><br><input type="button" style="font-size: 20px;" value="Add New Image" 
		onClick="location = 'addImage.asp?inspecID=<% = inspecID %>'; return false";></td></tr></table>
<% END IF	'--- noImages Check %>
</form>
<hr>
<% connSWPPP.Close 
Set connSWPPP = Nothing %>	
</body>
</html>