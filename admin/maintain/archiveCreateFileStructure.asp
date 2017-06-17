<%
Server.ScriptTimeout=7200 '7200 seconds or 2 hrs

If Not Session("validAdmin") Then
	Session("adminReturnTo") = Request.ServerVariables("path_info") & _
		"?" & Request.ServerVariables("query_string")
	Response.Redirect("loginUser.asp")
End If
DIM SID
SID= TRIM(Session("lastName"))
function validated(xValue, xType)
	validated=True
	SELECT CASE xType
		CASE "directory"
			SET re= NEW RegExp
			re.Global=true
			re.IgnoreCase=true
			re.pattern="([a-zA-Z]:\\([\w ]*)*\\)|([a-zA-Z]:\\)"
			IF NOT(re.Test(xValue)) THEN validated=False END IF
		CASE "integer"
			IF NOT(IsNumeric(xValue)) THEN validated=False END IF
	END SELECT
end function
'--IF NOT(validated(Request("destination"),"directory")) THEN Response.Redirect("archive.asp?err=001") END IF
IF NOT(validated(Request("projectID"),"integer")) THEN Response.Redirect("archive.asp?err=002") END IF
DIM fso, folder1, file1
SET fso = CreateObject("Scripting.FileSystemObject")
SQL0="sp_GetAllInspectionsforProject("& Request("projectID") &")" %>
<!-- #include file="../connSWPPP.asp" -->
<!-- #include file="archiveCreateFullReport.inc" -->
<!-- #include file="archiveCreateReportsHTML.inc" -->
<!-- #include file="archiveCreateAboutUs.inc" -->
<!-- #include file="archiveCreateContactUs.inc" -->
<!-- #include file="archiveCreateDefault.inc" --> <%
SET RS0=connSWPPP.execute(SQL0)
SQL2="SELECT projectName, projectPhase FROM Projects WHERE projectID="& Request("projectID") 
SET RS2= connSWPPP.execute(SQL2)
projectName=TRIM(RS2(0))
fullPName= TRIM(RS2(0)) &" "& TRIM(RS2(1))
SET RS2=nothing
'--- These next few lines create a psuedo-object type that allows us to interact with 
'--- and redim the array that we need to use for the file creation. Unlike normal VB 
'--- arrays we will leave the (0,0) slot blank. this will make it easier to read our
'--- loops later. This is also a 2 dimension array, but one of the axis' will be static
'--- while the other is dynamic.
Dim DynamicArray
'/// Declare constants
'-- declaring these constants just gives us a place to see what our
'-- array values are holding. Col1 is just a count of the other columns
Const Col1 = 5
Const col_arcID = 0			'-- this is an ID value (not really needed)
Const col_arcType=1			'-- this tells me what type of object we are going to create (directory, html file, image,.etc)
Const col_arcFilename = 2	'-- this gives me the name of the object
Const col_arcSrcPath = 3	'-- this gives me the source path to retrieve the object data from
Const col_arcSrcDest = 4 	'-- this gives me the full path destination to copy the object to (this is the local server location)
Const col_arcSrcClientDest=5'-- this gives me the full path destination to copy the object to (this would be the client location)
'-- time to instantiate the array to the default columns, this will also clear it.
ReDim DynamicArray(Col1, 0)
arr1 = DynamicArray
'-- This function is used to expand the scope of our dynamic array, and populate it with data.

Function AddArcElement(arcType, arcFileName, arcSrcPath, arcSrcDest, arcSrcClientDest) ' ///call function add slot
	ReDim Preserve arr1(Ubound(arr1,1), Ubound(arr1,2)+1)
	arr1(0,Ubound(arr1,2)) = Ubound(arr1,2)
	arr1(1,Ubound(arr1,2)) = arcType
	arr1(2,Ubound(arr1,2)) = arcFileName
	arr1(3,Ubound(arr1,2)) = arcSrcPath
	arr1(4,Ubound(arr1,2)) = arcSrcDest
	arr1(5,Ubound(arr1,2)) = arcSrcClientDest
End Function

'--- OK. It is time to populate the array. I will use the same code (mostly) 
'--- from the archivePreviewPage, I just won't display anything on the screen
'--- while I am populating the array.
'--- I also need to set the local dir on the server as a base path for the array and FSO's
imagePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & "images\"
localDest = Request.ServerVariables("APPL_PHYSICAL_PATH") & "admin\maintain\temporary_archives\"
clientDest= TRIM(Request("destination")) 
'--- Time to Create the Default Directory ---------------------------------------------------------------
	AddArcElement "dir","Root Folder","", localDest & SID, clientDest
IF (fso.FolderExists(localDest & SID)) THEN
	SET folder1 = fso.GetFolder(localDest & dirName)
ELSE
	SET folder1 = fso.CreateFolder(localDest & SID)
END IF
localDest = localDest & SID &"\"
DO WHILE NOT RS0.EOF
	dirName = DATEDIFF("d",#1/1/2000#,CDATE(RS0(1))) &"\"
'-- Now we check to see if this directory already exists. create it or get it ---------------------------
	AddArcElement "dir",dirName,dirName,localDest & dirName, clientDest & dirName
	IF (fso.FolderExists(localDest & dirName)) THEN
		SET folder1 = fso.GetFolder(localDest & dirName)
	ELSE
		SET folder1 = fso.CreateFolder(localDest & dirName)
	END IF
	AddArcElement "html","FullReport.html","database:inspecID="&RS0(0), localDest & dirName &"FullReport.html", clientDest & dirName &"FullReport.html"
'-- run the function to create the fullReport.asp file for this inspection report -----------------------
'-- we must pass the inspecID and the fullPath file information for the FSO.createFile ------------------
	createFullReportHTML RS0(0), localDest & dirName &"FullReport.html"
	SQL1="sp_GetOptionalImages("& RS0(0) &")" 
'-Response.Write(SQL1 &"<br>")
	SET RS1=connSWPPP.execute(SQL1)
	cnt1=1
	curOITDesc=""
   On Error Resume Next
	DO WHILE NOT RS1.EOF
		fileDesc= TRIM(RS1("oitDesc"))
		dir2Name=TRIM(RS1("oitName"))
		fileName= Trim(RS1("oImageFileName"))
		fileExt= "."& RIGHT(fileName,(LEN(fileName)-InStr(fileName,".")))
		if curOITDesc=fileDesc then 
			cnt1=cnt1+1
		else
			cnt1=1
			curOITDesc=fileDesc
		end if
		if (cnt1>1) then fileDesc=fileDesc &" "& cnt1 end if 
		AddArcElement fileExt, (fileDesc), (imagePath & dir2Name &"\"& fileName), (localDest & dirName & (fileDesc) & fileExt), (clientDest & dirName & (fileDesc) &"."& fileExt)
      If Err <> 0 Then
         Response.Write("Problem copying file " & arr1(3,UBound(arr1,2)))
      Else
		   fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
		End If
      RS1.MoveNext
	LOOP 
	AddArcElement "html","Default.html","database:inspecID="&RS0(0), localDest & dirName &"Default.html", clientDest & dirName &"Default.html"
	createDefault RS0(0), localDest & dirName 
'-- Cheack for any images associated with the Inspection Report -----------------------------
	imgSQLSELECT = "SELECT imageID, largeImage, smallImage, description FROM Images WHERE inspecID = " & RS0(0)
	Set rsImages = connSWPPP.execute(imgSQLSELECT)
	If Not rsImages.EOF Then
'--	create the sm and lg image file and array values ----------------------------------------
		AddArcElement "dir","sm","sm",localDest & dirName &"\sm", clientDest & dirName &"\sm"
		IF fso.FolderExists(localDest & dirName &"\sm") THEN fso.GetFolder(localDest & dirName &"\sm") ELSE fso.CreateFolder(localDest & dirName &"\sm") END IF
		AddArcElement "dir","lg","lg",localDest & dirName &"\lg", clientDest & dirName &"\lg"
		IF fso.FolderExists(localDest & dirName &"\lg") THEN fso.GetFolder(localDest & dirName &"\lg") ELSE fso.CreateFolder(localDest & dirName &"\lg") END IF
		Do While Not rsImages.EOF 
'--	create an array record of each image file and copy it to the correct dir ----------------
			AddArcElement "img", Trim(rsImages("smallImage")), imagePath &"\sm\"& Trim(rsImages("smallImage")), localDest & dirName & "\sm\"& Trim(rsImages("smallImage")), clientDest & dirName & "\sm\"& Trim(rsImages("smallImage"))
			fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
			AddArcElement "img", Trim(rsImages("largeImage")), imagePath &"\lg\"& Trim(rsImages("largeImage")), localDest & dirName & "\lg\"& Trim(rsImages("largeImage")), clientDest & dirName & "\lg\"& Trim(rsImages("largeImage"))
			fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
			RSImages.MoveNext
		LOOP
	End If
	RS0.MoveNext
LOOP
'-- now that all of the files are loaded. the last few things that need to happen are... ----
'--	create the autorun.inf, Reports.html, Images DIR and move over the default imgages ------
AddArcElement "inf","autorun.inf","", localDest & "autorun.inf", clientDest & "autorun.inf"
SET file1=fso.OpenTextFile(localDest &"autorun.inf",2,True)
file1.WriteLine("[autorun]")
file1.WriteLine("shellexecute=Reports.html")
file1.close 
AddArcElement "html","Reports.html","", localDest & "Reports.html", clientDest & "Reports.html"
createReportsHTML localDest, fullPName
AddArcElement "html","AboutUs.html","", localDest & "AboutUs.html", clientDest & "AboutUs.html"
createAboutUs localDest 
AddArcElement "html","ContactUs.html","", localDest & "ContactUs.html", clientDest & "ContactUs.html"
createContactUs localDest 
AddArcElement "html","global.css",Request.ServerVariables("APPL_PHYSICAL_PATH") & "\global.css", localDest & "global.css", clientDest & "global.css"
fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
AddArcElement "dir","Images","", localDest & "Images", clientDest & "Images"
IF fso.FolderExists(localDest & "\Images") THEN fso.GetFolder(localDest & "\Images") ELSE fso.CreateFolder(localDest & "\Images") END IF
AddArcElement "jpeg","logo.jpg", imagePath & "logo.jpg", localDest & "Images\swpppweb2.jpg", clientDest & "Images\swpppweb2.jpg"
fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
AddArcElement "jpeg","b&wlogoforreport.jpg", imagePath & "b&wlogoforreport.jpg", localDest & "Images\b&wlogoforreport.jpg", clientDest & "Images\b&wlogoforreport.jpg"
fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
AddArcElement "gif","acrobat.gif", imagePath & "acrobat.gif", localDest & "Images\acrobat.gif", clientDest & "Images\acrobat.gif"
fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
AddArcElement "gif","smallcamera.gif", imagePath & "smallcamera.gif", localDest & "Images\smallcamera.gif", clientDest & "Images\smallcamera.gif"
fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
'-- Now I have to move over all of the signature Files that match the inspections for this project --------
SQL2="SELECT DISTINCT u.signature FROM Users u, Inspections i WHERE i.userID=u.userID AND i.projectID="& Request("projectID")
SET RS2=connSWPPP.execute(SQL2)
'-- create the signatures dir and all of the signature files for the inspections --------------------------------------
AddArcElement "dir","signatures","", localDest & "Images\signatures", clientDest & "Images\signatures"
IF fso.FolderExists(localDest & "Images\signatures") THEN fso.GetFolder(localDest & "Images\signatures") ELSE fso.CreateFolder(localDest & "Images\signatures") END IF
DO WHILE NOT RS2.EOF
	AddArcElement "jpeg",TRIM(RS2(0)), imagePath &"signatures\"& TRIM(RS2(0)), localDest &"Images\signatures\"& TRIM(RS2(0)), clientDest &"Images\signatures\"& TRIM(RS2(0))
	fso.CopyFile arr1(3,UBound(arr1,2)), arr1(4,UBound(arr1,2)), True
	RS2.MoveNext
LOOP
SET RS2=nothing
'-- Now that I have the Report list for this project and an array to
'-- dump the result into it is time to create the files and get ready
'-- to copy them to the client machine.
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<head>
<title>SWPPP INSPECTIONS : Maintain : Archive : Create File Structure</title>
<link rel="stylesheet" href="../../global.css" type="text/css">
</head>
<!-- #include file="../adminHeader2.inc" -->
<table width="100%" border="0">
	<tr> 
		<td colspan=5><h1>SWPPP Inspections : Maintain : Archive : Create File Structure</h1></td>
	</tr>
	<tr><td colspan=5><h2>The File Structure has been created. Use your FTP tool to download the entire structure to a local directory on
		your harddrive. At that time you can can create CD-R(W)'s of this data for distribution.</td></tr>
</table>
</body>
</html>
<%	connSWPPP.Close()%>