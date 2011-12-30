<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Institution Information Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m012e0101.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">General Information</a> | </td>
		<td><a href="m012q0102.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">Funding Source</a></td>
	</tr>
</table>
</body>
</html>