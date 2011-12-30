<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Staff Information Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m002e0101.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>" target="SubBodyFrame">Personal Information</a> | </td>
		<td><a href="m002q0102.asp?insStaff_id=<%=Request.QueryString("insStaff_id")%>" target="SubBodyFrame">Employment Information</a></td>
	</tr>
</table>
</body>
</html>