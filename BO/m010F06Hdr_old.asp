<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Backorder Shipping Information Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m010e0601.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Method</a> |</td>
		<td><a href="m010e0602.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Address</a> |</td>
		<td><a href="m010e0603.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Schedule</a></td>
	</tr>
</table>
</body>
</html>