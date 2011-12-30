<!--------------------------------------------------------------------------
* File Name: m014FS3.asp
* Title: Purchase Requisition Number
* Main SP: 
* Description: This page is the frameset to display one purchase requisition.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!-- #Include File="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<title>PR Number: <%=ZeroPadFormat(Request.QueryString("insPurchase_Req_id"),8)%></title>
</head>
<frameset rows="*" cols="140,*" frameborder="0" framespacing="0"> 
  <frame name="MenuFrame" scrolling="NO" src="m014F3panel.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>">
  <frameset rows="18%,82%" cols="*" resize=no frameborder="NO" border="0" framespacing="0" > 
    <frame name="HeaderFrame" scrolling="NO" resize=no src="m014F3Hdr.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>">
    <frame name="BodyFrame" scrolling="YES" resize=no src="m014e0101.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>">
  </frameset>
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>

