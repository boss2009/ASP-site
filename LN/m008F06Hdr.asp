<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();
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
switch (String(rsLoan.Fields.Item("insEq_user_type").Value)){
	//client
	case "3":		
%>
		<td><a href="m008e0601.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="SubBodyFrame">Training Requested</a> |</td>
		<td><a href="m008e0602.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="SubBodyFrame">Training Status</a></td>	
<%
	break;
	//institution
	case "4":
%>
		<td><a href="m008e0601.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="SubBodyFrame">Training Requested</a> |</td>
		<td><a href="m008e0603.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" target="SubBodyFrame">Training Status</a></td>
<%	
	break;
	default:
%>
		<td>Training not available for this loan.</td>
<%
	break;
}
%>		
	</tr>
</table>
</body>
</html>