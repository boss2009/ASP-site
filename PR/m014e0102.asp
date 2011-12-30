<!--------------------------------------------------------------------------
* File Name: m014e0102.asp
* Title: Contract PO Number
* Main SP: cp_purchase_requisition_vendor
* Description: This page displays the Contract PO Number of a purchase requisition.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (Request.QueryString("Type")=="Standing") {
	var rsContractPO = Server.CreateObject("ADODB.Recordset");
	rsContractPO.ActiveConnection = MM_cnnASP02_STRING;
	rsContractPO.Source = "{call dbo.cp_Purchase_Requisition_Vendor("+ Request.QueryString("insPurchase_Req_id") + ",0)}";
	rsContractPO.CursorType = 0;
	rsContractPO.CursorLocation = 2;
	rsContractPO.LockType = 3;
	rsContractPO.Open();
}
%>
<html>
<head>
	<title>Contract PO Number</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<%
if (Request.QueryString("Type")=="Purchase") {
%>
<h5>Purchase Card Number</h5>
<%
Response.Write("<b>"+Request.QueryString("PurchaseCardNumber")+"/"+Request.QueryString("insPurchase_Req_id")+"</b>")
} else {
%>
<h5>Contract PO Number(s)</h5>
<hr>
<table cellspacing="1" cellpadding="2" class="Mtable">
	<tr> 		
		<th class="headrow" align="left">Vendor Name</th>
		<th class="headrow" align="left">Contract PO Number</th>	  
	</tr>
<% 
while (!rsContractPO.EOF) { 
%>
    <tr> 		
		<td><%=(rsContractPO.Fields.Item("chvVendorName").Value)%>&nbsp;</td>				
		<td><%=(rsContractPO.Fields.Item("chvContract_PO").Value)%>&nbsp;</td>
    </tr>
<%
	rsContractPO.MoveNext();
}
%>
</table>
<%
rsContractPO.Close();
}
%>
</body>
</html>