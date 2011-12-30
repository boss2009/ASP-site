<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
<title>Shipping Information Frame Set</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0">
  <frame name="SubMenuFrame" scrolling="NO"  noresize src="m009F04Hdr.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_dtl_id=<%=Request.QueryString("intShip_dtl_id")%>">
  <frame name="SubBodyFrame" scrolling="yes" src="m009e0401.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_dtl_id=<%=Request.QueryString("intShip_dtl_id")%>">
</frameset>
<noframes>
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes>
</html>

