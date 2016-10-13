/*
** Code Name: JavaScript1.2 Form Validation
** Developer: Plonka Interactive - David K. Watts
** Date/Time: Friday, July 30, 2002 - 8:24:00 a
** Desc/Info: Form validation for Local
**  Hair Stylist customers.
** FileLocal: root/admin/
** Parameter: Use JavaScript header src
*/
/* Not used this time!

function isEmail(string) {
	if (string.search(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/) != -1) {
		return true;
	} else {
		return false;
	}
}
End Not used this time! */
function isProper(string) {
	if (string.search(/^\w+( \w+)?$/) != -1) {
		return true;
	} else {
		return false;
	}
}
