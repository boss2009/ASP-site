<!--------------------------------------------------------------------------
* File Name: m014F3Hdr.asp
* Title: Purchase Requisition Header
* Main SP: cp_frmhdr
* Description: This page displays the header information of a purchase
* requisition.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsPurchaseHeader = Server.CreateObject("ADODB.Recordset");
rsPurchaseHeader.ActiveConnection = MM_cnnASP02_STRING;
rsPurchaseHeader.Source = "{call dbo.cp_FrmHdr(14,"+Request.QueryString("insPurchase_Req_id")+")}";
rsPurchaseHeader.CursorType = 0;
rsPurchaseHeader.CursorLocation = 2;
rsPurchaseHeader.LockType = 3;
rsPurchaseHeader.Open();
%>
<html>
<head>
	<title>Purchase Requisition Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="height:60px"> 
<table cellspacing="1" cellpadding="1" border="0">
	<tr> 
		<td nowrap valign="top" width="110"><b>PR Number:</b></td>
		<td nowrap valign="top" width="180"><%=ZeroPadFormat(Request.QueryString("insPurchase_Req_id"),8)%></td>
		<td nowrap valign="top" width="110"><b>Date Ordered:</b></td>
		<td nowrap valign="top" width="130"><%=FilterDate(rsPurchaseHeader.Fields.Item("dtsDate_Ordered").Value)%></td>
    </tr>
    <tr> 
		<td nowrap valign="top"><b>Vendor:</b></td>
		<td nowrap valign="top" colspan="3"><%=rsPurchaseHeader.Fields.Item("chvVendor").Value%></td>
    </tr>
</table>
</div>
</body>
</html>
<%
rsPurchaseHeader.Close();
%>