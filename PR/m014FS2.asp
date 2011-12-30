<!--------------------------------------------------------------------------
* File Name: m014FS2.asp
* Title: Purchase Requisition Top Menu
* Main SP: 
* Description: This page is the parent frameset for purchase requisition module.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Purchase Requisition Top Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="140,*" frameborder="No" framespacing="0" rows="*"> 
	<frame name="PurchaseBrowseLeftFrame" scrolling="NO" noresize src="m014F2Hdr.asp">
	<frame name="PurchaseBrowseRightFrame" scrolling="YES" noresize src="m014s0101.asp">
</frameset>
<noframes> 
<body>
Your browser either has frame disabled or does not support frames.
</body>
</noframes> 
</html>

