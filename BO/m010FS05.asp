<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
<title>Shipping Information Frame Set</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m010F05Hdr.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>">
	<frame name="SubBodyFrame" scrolling="yes" src="m010e0502.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
</html>
<a href="m010e0502.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>" target="SubBodyFrame">Address</a> 