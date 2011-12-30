<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Search Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr> 
<!--	<td><a href="m008q01lw.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>" target="SearchPageBody">By Tree Browse</a> |</td>-->
		<td><a href="m010p0101.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>" target="SearchPageBody">Search Class</a> |</td>
		<td><a href="m010p0102.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>" target="SearchPageBody">Search Bundle</a></td>		
	</tr>
</table>
</body>
</html>