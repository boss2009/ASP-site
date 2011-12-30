<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Accessories & Notes Frame Set</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m010F04Hdr.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>" >
	<frame name="SubBodyFrame" scrolling="yes" src="m010q0402.asp?intBuyout_Req_id=<%=Request.QueryString("intBuyout_Req_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
</html>
