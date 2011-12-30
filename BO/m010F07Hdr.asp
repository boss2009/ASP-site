<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();
%>
<html>
<head>
	<title>Training Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table cellpadding="1" cellspacing="1">
	<tr>     	
<%
switch (String(rsBuyout.Fields.Item("insEq_user_type").Value)){
	//client
	case "3":		
%>
		<td><a href="m010e0701.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Training Requested</a> |</td>
		<td><a href="m010e0702.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Training Status</a></td>	
<%
	break;
	//institution
	case "4":
%>
		<td><a href="m010e0701.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Training Requested</a> |</td>
		<td><a href="m010e0703.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Training Status</a></td>
<%	
	break;
	default:
%>
		<td>Training not available for this buyout.</td>
<%
	break;
}
%>		
	</tr>
</table>
</body>
</html>