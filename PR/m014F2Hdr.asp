<!--------------------------------------------------------------------------
* File Name: m014F2Hdr.asp
* Title: Purchase Requisition Menu
* Main SP: 
* Description: This page displays functions of purchase requisition module.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<html>
<head>
	<title>Purchase Requisition Menu</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/PanelStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>		
</head>
<body>
<table align="center" cellspacing="0">
	<tr>
		<td align="center"><div align="center"><a href="javascript: top.window.close();"><img src="../i/tn_pur_req_02.jpg" ALT="Return to Main Menu." width="80" height="60"></a></div></td>
	</tr>
	<tr> 
		<td height="18px" class="MenuItem" align="center" width="120"><a href="m014d01.asp" target="PurchaseBrowseRightFrame">Desktop</a></td>
	</tr>
	<tr> 		
	    <td height="18px" class="MenuItem" align="center"><a href="m014q01.asp" target="PurchaseBrowseRightFrame">Browse All</a></td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m014s0101.asp" target="PurchaseBrowseRightFrame">Quick Search</a></td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m014s0102.asp" target="PurchaseBrowseRightFrame">Advanced Search</a></td>
	</tr>
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m014s0103.asp" target="PurchaseBrowseRightFrame">Power Search</a></td>
	</tr>	
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="m014r0101.asp" target="PurchaseBrowseRightFrame">Reports</a></td>
	</tr>	
	<tr>
		<td height="18px" class="MenuItem" align="center"><a href="javascript: openWindow('m014a0101.asp','wQA14');">New Purchase Requisition</a></td>
	</tr>	
</table>
</body>
</html>