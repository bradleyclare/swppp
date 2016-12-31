<table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr bgcolor="#e5e6e8">	
  <td class="header" valign="bottom" align="left" colspan=2><img 
  	src="../images/swpppweb2.jpg" height="101" border="0" vspace="0" hspace="0" 
		alt="SWPPP Inspections"></td>
  <td align="center" valign="middle" width="100%" nowrap><p class="main">Website Administration</p></td></tr>
<tr bgcolor="#999999"><td colspan="3">&nbsp;  </td></tr>
</table><table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr><!-- Start Navigation -->
	<td height="600" width=125 valign="top" align="right" bgcolor="#CCCCCC" nowrap><br>
		
		<% If Session("validAdmin") or Session("validInspector") then %>
			<a href="maintain/viewReports.asp" class="admin">View Reports</a><br>
			<a href="maintain/addReport.asp" class="admin">Add Report</a><br><br>
			<a href="maintain/releaseReports.asp" class="admin">Release Reports</a><br>
      <a href="maintain/releaseRepeatItems.asp" class="admin">Repeat Item Alert</a><br>
      <a href="maintain/releaseOpenItems.asp" class="admin">Open Item Alert</a><br>
      <a href="maintain/manageGroups.asp" class="admin">Manage Groups</a><br><br>
		<% End If %>
		<% If Session("validAdmin") then %>
	      	<a href="maintain/viewUsersAdmin.asp" class="admin">View Users</a><br>
		<% end if %>			
		<% If Session("validDirector") then %>
	      	<a href="maintain/viewUsersDir.asp" class="admin">View Users</a><br>
		<% end if %>			
		<% If Session("validAdmin") or Session("validDirector") then %>
			<a href="maintain/addUser.asp" class="admin">Add User</a><br>
			<a href="maintain/reportUsersByProject.asp" class="admin">Users by Project</a><br><br>
		<% end if %>		
		<% If Session("validAdmin") then 
	testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
	If InStr(testStr,Session("email"))>0 Then %>
			<a href="maintain/viewProjects.asp" class="admin">View Projects</a><br>
			<a href="maintain/commissions.asp" class="admin">Commissions</a><br><br>
<%	End If %>
			<a href="maintain/adminAbout.asp" class="admin">Modify About Us</a><br>
	    	<a href="maintain/adminContact.asp" class="admin">Modify Contact Us</a><br><br>
	    	Remove Files<br>
	    	<a href="maintain/remImgFiles.asp" class="admin">Images</a><br>
	    	<a href="maintain/remSigFiles.asp" class="admin">Signatures</a><br><br>
		<% end if %>
		
		<% If Session("validAdmin") then %>
		<a href="maintain/archive.asp" class="admin">Archive Reports</a><br><br>
		<% End IF %>
		<a href="../" class="admin">Public Home</a><br><br>
    <a href="../logout.asp" class="admin">Logout</a><br/><br/>
</td>
<!-- End Navigation --><td colspan="2" valign="top" width="90%"><br>