<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<html>
<head>
<title>Inventory Class Search</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="30,*" frameborder="NO" border="0" framespacing="0" cols="*"> 
  <frame name="SearchPageHeader" scrolling="NO" noresize src="m008p01Hdr.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>" >
  <frame name="SearchPageBody" src="m008p0101.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intEqpReq_Id=<%=Request.QueryString("intEqpReq_Id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>
