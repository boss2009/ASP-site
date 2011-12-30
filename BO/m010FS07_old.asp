<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
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
<title>Training Frame</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
switch (String(rsBuyout.Fields.Item("insEq_user_type").Value)){
	//client
	case "3":
%>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0">
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m010F07Hdr.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" >
	<frame name="SubBodyFrame" scrolling="yes" src="m010e0701.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>">
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
<%
	break;
	//institution
	case "4":
%>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0">
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m010F07Hdr.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" >
	<frame name="SubBodyFrame" scrolling="yes" src="m010e0701.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>">
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
<%
	break;
	default:
%>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0">
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m010F07Hdr.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" >
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
<%
	break;
}
%>
</html>