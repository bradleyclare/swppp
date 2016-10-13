/*
** Code Name: JavaScript1.2 Form Validation
** Developer: Plonka Interactive - David K. Watts
** Date/Time: Tuesday, August 12, 2002 - 7:24:00 PM
** Desc/Info: Form validation for Web
**  Clients.
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
