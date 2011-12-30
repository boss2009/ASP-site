<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

if (rsBuyout.Fields.Item("insEq_user_type").Value=="3") {
	Response.Redirect("../AC/m001a0902.asp?intAdult_id="+rsBuyout.Fields.Item("intEq_user_id").Value+"&intBuyout_Req_id="+Request.QueryString("intBuyout_Req_id"));
}
if (rsBuyout.Fields.Item("insEq_user_type").Value=="4") {
	Response.Redirect("../SH/m012a0802.asp?Type=Buyout&insSchool_id="+rsBuyout.Fields.Item("intEq_user_id").Value+"&intBuyout_Req_id="+Request.QueryString("intBuyout_Req_id"));
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
<i>Correspondence not available for this buyout.</i>
</body>
</html>