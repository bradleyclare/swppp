/*
** Code Name: JavaScript Form Validation
** Developer: Plonka Interactive - David K. Watts
** Date/Time: Monday, July 30, 2002 - 8:24:00 a
** Desc/Info: Form validation for Local
**  Hair Stylist customers.
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
	if (form.localFile.value == "") {
		alert("Please browse for a file.");
		form.localFile.focus();
		return false;
	}
	return true;
}