/*
** Code Name: JavaScript Form Validation
** Developer: Plonka Interactive - David K. Watts
** Date/Time: Friday, July 26, 2002 - 4:24:00 PM
** Desc/Info: Login form validation for
**  Local Hair Stylist customers.
** FileLocal: root/admin/
** Parameter: Use JavaScript header src
*/
function isEmail(string) {
	if (!string) return false;
	var iChars = "*|,\":<>[]{}`\';()&$#%";

	for (var i = 0; i < string.length; i++) {
		if (iChars.indexOf(string.charAt(i)) != -1)
			return false;
		}
	return true;
}
                      
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
	/* Not used this time!
	// Added for blank input!
	if (form.name.value == "") {
		alert("Please enter a valid name.");
		form.name.focus();
		return false;
	}
	End Not used this time! */
	return true;
}