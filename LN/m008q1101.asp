<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

if (rsLoan.Fields.Item("insEq_user_type").value=="3") {
	Response.Redirect("../AC/m001a0903.asp?intLoan_Req_id="+Request.QueryString("intLoan_Req_id")+"&intAdult_id="+rsLoan.Fields.Item("intEq_user_id").Value);
}
if (rsLoan.Fields.Item("insEq_user_type").Value=="4") {
	Response.Redirect("../SH/m012a0802.asp?Type=Loan&intLoan_Req_id="+Request.QueryString("intLoan_Req_id")+"&insSchool_id="+rsLoan.Fields.Item("intEq_user_id").Value);
}
%>
<html>
<head>
	<title>Correspondence</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Correspondence</h5>
<hr>
<i>Correspondence unavailable for staff user.</i>
</body>
</html>