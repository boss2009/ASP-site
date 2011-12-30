<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>PILAT Referral Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table width="381" cellpadding="1" cellspacing="1">
	<tr> 
    	<td><a href="m012e0201.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">Temp Referral Type</a> |</td>
		<td><a href="m012e0202.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">Referral Details</a> |</td>
<!--	<td><a href="m012q0203.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">Equipment Requested</a> |</td>-->
		<td><a href="m012q0204.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">On-site Support</a> |</td>
		<td><a href="m012q0205.asp?intReferral_id=<%=Request.QueryString("intReferral_id")%>&insSchool_id=<%=Request.QueryString("insSchool_id")%>" target="SubBodyFrame">Temp Student</a></td>		
	</tr>
</table>
</body>
</html>