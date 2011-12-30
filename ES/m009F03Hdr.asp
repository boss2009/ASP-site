<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Service Performed Frame Panel</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m009e0301.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="SubBodyFrame">In Service</a> | </td>
		<td><a href="m009e0302.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" target="SubBodyFrame">Out Service</a></td>
	</tr>
</table>
</body>
</html>