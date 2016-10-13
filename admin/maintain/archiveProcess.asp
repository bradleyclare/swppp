<% '--- Instantiate the FileUp object
Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")

'--- Assign the same progress ID that we assigned to the progress object
oFileUp.ProgressID = CInt(Request.QueryString("progressid"))
	
oFileUp.Path = Server.MapPath(Application("vroot") & "/temp")
...
oFileUp.Form("myFile").Save %>