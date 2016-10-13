/*
** Code Name: JavaScript Form Validation
** Developer: Plonka Interactive - David K. Watts
** Date/Time: Tuesday, August 12, 2002 - 7:24:00 PM
** Desc/Info: Form validation for Web
**  Clients.
** FileLocal: root/admin/
** Parameter: Use JavaScript header src
*/

/* Not used this time!
function isEmail(string) {
	if (!string) return false;
	var iChars = "*|,\":<>[]{}`\';()&$#%";

	for (var i = 0; i < string.length; i++) {
		if (iChars.indexOf(string.charAt(i)) != -1)
			return false;
		}
	return true;
}
End Not used this time! */

function isProper(string) {
	if (!string) return false;
	var iChars = "*|,\":<>[]{}`\';()@&$#%";

	for (var i = 0; i < string.length; i++) {
		if (iChars.indexOf(string.charAt(i)) != -1)
		return false;
	}
	return true;
}
function isReady(form) {
	/* Not used this time!
	if (isEmail(form.email.value) == false) {
		alert("Please enter a valid email address.");
		form.email.focus();
		return false;
	}
	if (isProper(form.pswrd.value) == false) {
		alert("Please enter a valid password.");
		form.pswrd.focus();
		return false;
	}
	End Not used this time! */
	
	// Added for blank input!
	if (form.projectName.value == "") {
		alert("Please enter a valid name.");
		form.projectName.focus();
		return false;
	}
	return true;
}