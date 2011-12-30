<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<html>
<head>
<title>Inventory Class Search</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="30,*" frameborder="NO" border="0" framespacing="0" cols="*"> 
  <frame name="SearchPageHeader" scrolling="NO" noresize src="m010p01Hdr.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>" >
  <frame name="SearchPageBody" src="m010p0101.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&insBO_Eqp_Rqst_id=<%=Request.QueryString("insBO_Eqp_Rqst_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames. 
</body>
</noframes> 
</html>
