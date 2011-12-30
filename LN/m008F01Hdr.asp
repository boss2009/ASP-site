<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Loan Request Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body cellpadding="1" cellspacing="1">
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m008e0101.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="SubBodyFrame">Loan Request</a> | </td>
		<td><a href="m008e0102.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="SubBodyFrame">TAP Date</a></td>
	</tr>
</table>
</body>
</html>