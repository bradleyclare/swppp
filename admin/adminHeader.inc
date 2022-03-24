<table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr bgcolor="#e5e6e8">	
  <td class="header" valign="bottom" align="left" colspan=2><img 
  	src="../images/swpppweb2.jpg" height="101" border="0" vspace="0" hspace="0" 
		alt="SWPPP Inspections"></td>
  <td align="center" valign="middle" width="100%" nowrap><p class="main">website administration</p></td></tr>
<tr bgcolor="#999999"><td colspan="3">&nbsp;  </td></tr>
</table><table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr><!-- Start Navigation -->
	<td height="600" width=125 valign="top" align="right" bgcolor="#CCCCCC" nowrap><br>
		
		<% If Session("validAdmin") or Session("validInspector") then %>
			<a href="maintain/viewReports.asp" class="admin">reports</a><br>
			<a href="maintain/addReport.asp" class="admin">add report</a><br><br>
			<a href="maintain/releaseReports.asp" class="admin">email reports</a><br>
      		<a href="maintain/releaseRepeatItems.asp" class="admin">repeat item alert</a><br>
      <% End If %>
		<% If Session("validAdmin") then %>
			<a href="maintain/releaseOpenItems.asp" class="admin">open item alert</a><br><br>
		<% else %>
			<br>
		<% end if %>
		<% If Session("validAdmin") or Session("validInspector") then %>
			<a href="maintain/recentComments.asp" class="admin">notes</a><br><br>
		<% End If %>
		<% If Session("validAdmin") then %>
			<a href="maintain/viewUsersAdmin.asp" class="admin">users</a><br>
		<% end if %>			
		<% If Session("validDirector") then %>
	    	<a href="maintain/viewUsersDir.asp" class="admin">users</a><br>
		<% end if %>			

		<% If Session("validAdmin") or Session("validDirector") then %>
			<a href="maintain/addUser.asp" class="admin">add user</a><br>
			<a href="maintain/reportUsersByProject.asp" class="admin">users by project</a><br>
         	<a href="maintain/reportProjectsByUser.asp" class="admin">projects by user</a><br><br>
		<% end if %>		
		<% If Session("validAdmin") then %> 
			<a href="maintain/manageGroups.asp" class="admin">customers</a><br>
	   		<a href="maintain/manageUserGroups.asp" class="admin">user groups</a><br>
      		<a href="maintain/manageCommonItems.asp" class="admin">items</a><br><br>
			<% testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
			If InStr(testStr,Session("email"))>0 Then %>
				<a href="maintain/viewProjects.asp" class="admin">projects</a><br><br>
			<%	End If %>
			<a href="maintain/adminAbout.asp" class="admin">modify about us</a><br>
	    	<a href="maintain/adminContact.asp" class="admin">modify contact us</a><br><br>
		<% end if %>
		
		<% If Session("validAdmin") then %>
			<a href="maintain/cleanupScripts.asp" class="admin">cleanup scripts</a><br>
			<!--<a href="maintain/archive.asp" class="admin">archive reports</a><br><br>-->
		<% End IF %>
		<br/>
		<a href="../views/projects.asp" class="admin">public</a><br><br>
    <a href="../logout.asp" class="admin">logout</a><br/><br/>
</td>
<!-- End Navigation --><td colspan="2" valign="top" width="90%"><br>