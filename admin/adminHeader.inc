<div class="header-top">
<div class="fl">
	<img src="../images/swpppweb2.jpg" height="100" border="0" vspace="0" hspace="0" alt="SWPPP Inspections">
</div>
<h2>Website Administration</h2>
</div>
<div class="three columns alpha nav-side">
		<% If Session("validAdmin") or Session("validInspector") then %>
			<a href="maintain/viewReports.asp" class="admin">View Reports</a>
			<a href="maintain/addReport.asp" class="admin">Add Report</a>
		<% end if %>
		<% If Session("validInspector") Then %>
			<a href="maintain/releaseReports.asp" class="admin">Release Reports</a>
		<% End If %>
		<% If Session("validAdmin") then %>
	      	<a href="maintain/viewUsersAdmin.asp" class="admin">View Users</a>
		<% end if %>			
		<% If Session("validDirector") then %>
	      	<a href="maintain/viewUsersDir.asp" class="admin">View Users</a>
		<% end if %>			
		<% If Session("validAdmin") or Session("validDirector") then %>
			<a href="maintain/addUser.asp" class="admin">Add User</a>
			<a href="maintain/reportUsersByProject.asp" class="admin">Users by Project</a>
		<% end if %>		
		<% If Session("validAdmin") then 
	testStr="dwims@swpppinspections.com:jwright@swpppinspections.com"
	If InStr(testStr,Session("email"))>0 Then %>
			<a href="maintain/viewProjects.asp" class="admin">View Projects</a>
			<a href="maintain/commissions.asp" class="admin">Commissions</a>
<%	End If %>
			<a href="maintain/adminAbout.asp" class="admin">Modify About Us</a>
	    	<a href="maintain/adminContact.asp" class="admin">Modify Contact Us</a>
	    	Remove Files<br>
	    	<a href="maintain/remImgFiles.asp" class="admin">Images</a>
	    	<a href="maintain/remSigFiles.asp" class="admin">Signatures</a>
		<% end if %>
		
		<% If Session("validAdmin") then %>
		<a href="maintain/archive.asp" class="admin">Archive Reports</a>
		<% End IF %>
		<a href="../" class="admin">Public Home</a>
</div> <!-- End Navigation -->
<div class="nine columns omega content">