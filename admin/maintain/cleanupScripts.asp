<%
If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info")
	Response.Redirect("loginUser.asp")
End If
%>

<html>
<head>
	<title>SWPPP INSPECTIONS :: Cleanup Scripts </title>
	<link rel="stylesheet" type="text/css" href="../../global.css">
</head>
<body>
<!-- #include file="../adminHeader2.inc" -->
<h1>Cleanup Scripts</h1>
<h3><a href="testscripts/changeInvoiceMemo.asp">Cleanup Invoice Memos</a></h3>
<p>Fix any memos that say "Thank you for your business." to "Due upon receipt."</p>
<h3><a href="testscripts/cleanupAddresses.asp">Cleanup Addresses</a></h3>
<p>Attempts to remove any trailing commas and spaces in the address name</p>
<h3><a href="testscripts/cleanupApprovalDates.asp">Cleanup Approval Dates</a></h3>
<p>This will cleanup any incorrect approval dates that were set to 1900.  It will be set to the previous date.</p>
<h3><a href="testscripts/cleanupComments.asp">Change Comments</a></h3>
<p>This will update the project id for any comment that does not already have one set.</p>
<h3><a href="testscripts/cleanupScoring.asp">Cleanup Scoring</a></h3>
<p>This will check and fix any scoring counts that do not add up. Like completed items more than total items.  This takes a very long time to run.</p>
<h3><a href="testscripts/cleanupUserRights.asp">Cleanup User Rights</a></h3>
<p>This script will make sure the highest rights field corresponds correctly to all the rights on all the projects for each user.</p>
</div>
</td></tr></table>
</body>
</html>