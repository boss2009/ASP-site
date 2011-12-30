<!--------------------------------------------------------------------------
* File Name: m014p01FS.asp
* Title: Inventory Class Search
* Main SP: 
* Description: This page is the frameset for inventory class search.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<html>
<head>
	<title>Inventory Class Search</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="30,*" frameborder="NO" border="0" framespacing="0" cols="*"> 
	<frame name="SearchPageHeader" scrolling="NO" noresize src="m014p01Hdr.asp?insPurchase_req_id=<%=Request.QueryString("insPurchase_req_id")%>&insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>" >
	<frame name="SearchPageBody" src="m014p0101.asp?insPurchase_req_id=<%=Request.QueryString("insPurchase_req_id")%>&insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>
