<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Client Information Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m001e0101.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="SubBodyFrame">General Information</a> | </td>
		<td><a href="m001q0102.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="SubBodyFrame">Employment Information</a></td>
	</tr>
</table>
</body>
</html>