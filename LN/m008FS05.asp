<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
<title>Shipping Information Frame Set</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
	<frame name="SubMenuFrame" scrolling="NO"  noresize src="m008F05Hdr.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>" >
	<frame name="SubBodyFrame" scrolling="yes" src="m008e0501.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
</html>
