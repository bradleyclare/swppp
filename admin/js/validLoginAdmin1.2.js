/*
** Code Name: JavaScript1.2 Form Validation
** Developer: Plonka Interactive - David K. Watts
** Date/Time: Friday, July 26, 2002 - 4:24:00 PM
** Desc/Info: Login form validation for
**  Local Hair Stylist customers.
** FileLocal: root/admin/
** Parameter: Use JavaScript header src
*/
function isEmail(string) {
	if (string.search(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/) != -1) {
		return true;
	} else {
		return false;
	}
}
/* Not used this time!
function isProper(string) {
	if (string.search(/^\w+( \w+)?$/) != -1) {
		return true;
	} else {
		return false;
	}
}
End Not used this time! */