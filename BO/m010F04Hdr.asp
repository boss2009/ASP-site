<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Accessories & Notes Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><a href="m010q0402.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" target="SubBodyFrame">Equipment Requested Notes</a> |</td>
		<td><a href="m010q0403.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" target="SubBodyFrame">Equipment Sold Notes</a> |</td>
    	<td><a href="m010q0401.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" target="SubBodyFrame">Accessories</a></td>		
	</tr>
</table>
</body>
</html>