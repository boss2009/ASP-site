<!--------------------------------------------------------------------------
* File Name: m014p01Hdr.asp
* Title: Search Frame Header
* Main SP: 
* Description: This page is the search frame header.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Search Frame Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<table width="280">
	<tr> 
<!--	<td><a href="m014q01lw.asp?insPurchase_req_id=<%=Request.QueryString("insPurchase_req_id")%>&insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>" target="SearchPageBody">Search in Tree View</a></td>-->
		<td><a href="m014p0101.asp?insPurchase_req_id=<%=Request.QueryString("insPurchase_req_id")%>&insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>" target="SearchPageBody">Search by Class Name</a></td>
	</tr>
</table>
</body>
</html>