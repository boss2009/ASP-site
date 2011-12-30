<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<html>
<head>
	<title>Forms and Reports</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	</Script>
</head>
<body>
<h5>Forms and Reports</h5>
<hr>
<a href="LoanPackingList.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>" target="_blank">Loan Packing List</a>
</body>
</html>