<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Buyout Request Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m010e0101.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" target="SubBodyFrame">Buyout Request</a> | </td>
		<td><a href="m010e0102.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" target="SubBodyFrame">TAP Date</a> | </td>
		<td><a href="m010q0103.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" target="SubBodyFrame">Funding Source</a></td>
	</tr>
</table>
</body>
</html>