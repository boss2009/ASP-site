<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
<%
var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();
%>
<html>
<head>
	<title>Buyout No: <%=ZeroPadFormat(Request.QueryString("intBuyout_Req_id"),8)%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="*" cols="147,*" frameborder="0" framespacing="0">
	<frame name="MenuFrame" scrolling="NO" src="m010F3panel.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>">
  <frameset rows="25%,77%" cols="*" frameborder="NO" border="0" framespacing="0" >
    <frame name="HeaderFrame" scrolling="NO"  src="m010F3Hdr.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>">
    <frame name="BodyFrame" scrolling="YES" src="m010FS01.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>">
  </frameset>
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>
<%
rsBuyout.Close();
%>