<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>General Information Frame Set</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="no" border="0" framespacing="0"> 
  <frame name="SubMenuFrame" scrolling="no" noresize  src="m009F01Hdr.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" >
  <frame name="SubMenuFrame" scrolling="no" noresize  src="m009e0101.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>" >
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>
