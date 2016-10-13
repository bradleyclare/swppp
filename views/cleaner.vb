<%
function CleanText(textStr)
	CleanText=REPLACE(textStr,"/*","")
	CleanText=REPLACE(CleanText,"*/","")
	CleanText=REPLACE(CleanText,chr(34),"&quot;")
	CleanText=REPLACE(CleanText,chr(39),"&apos;")
	CleanText=REPLACE(CleanText,chr(45),"&hyphen;")
	CleanText=REPLACE(CleanText,chr(145),"&lsquo;")
	CleanText=REPLACE(CleanText,chr(147),"&ldquo;")
	CleanText=REPLACE(CleanText,chr(169),"&copy;")
	CleanText=REPLACE(CleanText,chr(174),"&reg;")
end function 
function UnCleanText(textStr)
	UnCleanText=REPLACE(textStr,"&quot;",chr(34))
	UnCleanText=REPLACE(UnCleanText,"&apos;",chr(39))
	UnCleanText=REPLACE(UnCleanText,"&hyphen;",chr(45))
	UnCleanText=REPLACE(UnCleanText,"&lsquo;",chr(145))
	UnCleanText=REPLACE(UnCleanText,"&ldquo;",chr(147))
	UnCleanText=REPLACE(UnCleanText,"&copy;",chr(169))
	UnCleanText=REPLACE(UnCleanText,"&reg;",chr(174))
end function
%>