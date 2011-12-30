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
<!--	<td><a href="m008q01lw.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>" target="SearchPageBody">By Tree Browse</a></td>-->
		<td><a href="m008p0101.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>" target="SearchPageBody">Search Class</a> |</td>
		<td><a href="m008p0102.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>" target="SearchPageBody">Search Bundle</a></td>		
	</tr>
</table>
</body>
</html>