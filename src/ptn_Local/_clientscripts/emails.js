//<script language="javascript">

// Validates email addresses
function validEMail(a)
{
	var pos1 = a.lastIndexOf('@');
	var pos2 = a.indexOf('.');

	return ( (pos1 != -1) && (pos2 != -1) && (pos1 != 0) && (pos1 != a.length) && 
		(pos2 != 0) && (pos2 != a.length) && (pos1 == a.indexOf('@')) && 
		(a.charAt(pos1-1) != ' ') && (a.charAt(pos1+1) != ' ') && 
		(a.charAt(pos2-1) != ' ') && (a.charAt(pos2+1) != ' ') );
}
