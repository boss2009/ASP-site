<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Follow-Up Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><a href="m001q1101.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="FollowUpFrameBody">Annual Follow-Up</a> | </td>
		<td><a href="m001q1102.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="FollowUpFrameBody">EPPD Buyout Follow-Up</a> | </td>
		<td><a href="m001q1103.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>" target="FollowUpFrameBody">General Follow-Up</a></td>
	</tr>
</table>
</body>
</html>