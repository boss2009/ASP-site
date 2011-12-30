<!--------------------------------------------------------------------------
* File Name: m014q0102.asp
* Title: Purchase Requisition
* Main SP: cp_purchase_requisition_vendor
* Description: This page lists vendors and Contract PO Numbers..
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" --> 
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsVendorContractPO = Server.CreateObject("ADODB.Recordset");
rsVendorContractPO.ActiveConnection = MM_cnnASP02_STRING;
rsVendorContractPO.Source = "{call dbo.cp_Purchase_Requisition_Vendor("+ Request.QueryString("insPurchase_Req_id") + ",0)}";
rsVendorContractPO.CursorType = 0;
rsVendorContractPO.CursorLocation = 2;
rsVendorContractPO.LockType = 3;
rsVendorContractPO.Open();
%>
<html>
<head>
	<title>Purchase Requisition: <%=Request.QueryString("insPurchase_Req_id")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Vendors and Contract PO Numbers</h5>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
	<tr> 
		<th nowrap class="headrow" align="left">Vendor</th>
		<th nowrap class="headrow" align="left">Contract PO Number</th>
	</tr>
<% 
while (!rsInventoryRequested.EOF) { 
%>
    <tr> 
		<td nowrap><%=(rsVendorContractPO.Fields.Item("chvVendorName").Value)%></td>
		<td nowrap><%=(rsVendorContractPO.Fields.Item("chvContract_PO").Value)%></td>
    </tr>
<%
	rsVendorContractPO.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsVendorContractPO.Close();
%>