<!--------------------------------------------------------------------------
* File Name: m014q0501.asp
* Title: Backorder Received Dates
* Main SP: cp_PR_BackOrder_Rx
* Description: This page lists the backorder received dates of a purchase
* requisition.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsBackOrders = Server.CreateObject("ADODB.Recordset");
rsBackOrders.ActiveConnection = MM_cnnASP02_STRING;
rsBackOrders.Source = "{call dbo.cp_PR_BackOrder_Rx(0,"+Request.QueryString("insPurchase_Req_id")+",'',0,0,0,'Q',0,0)}";
rsBackOrders.CursorType = 0;
rsBackOrders.CursorLocation = 2;
rsBackOrders.LockType = 3;
rsBackOrders.Open();

var total = 0;
while (!rsBackOrders.EOF) {
	total++;
	rsBackOrders.MoveNext();
}
rsBackOrders.Requery();
%>
<html>
<head>
	<title>Backorder Received</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<Script language="Javascript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}
	</Script>
</head>
<body>
<h5>Backorder Received</h5>
<table cellspacing="1">
	<tr> 
		<td align="left">Displaying <b><%=(total)%></b> Records</td>
	</tr>
</table>
<hr>
<table cellpadding="2" cellspacing="1" class="Mtable">
    <tr> 
		<th nowrap class="headrow" align="left">Date Received</th>
		<th nowrap class="headrow" align="left">Inventory Class</th>
		<th nowrap class="headrow" align="left">Quantity</th>	 
	</tr>
<% 
while (!rsBackOrders.EOF) { 
%>
    <tr> 
		<td nowrap><a href="m014e0501.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>&intBack_Ord_id=<%=(rsBackOrders.Fields.Item("intBack_Ord_id").Value)%>"><%=FilterDate(rsBackOrders.Fields.Item("dtsBack_Ord_rx").Value)%>&nbsp;</a></td>
		<td nowrap><%=(rsBackOrders.Fields.Item("chvEqp_Class").Value)%></td>
		<td nowrap align="center"><%=(rsBackOrders.Fields.Item("intQuantity").Value)%></td>
    </tr>
<%
	rsBackOrders.MoveNext();
}
%>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><a href="javascript: openWindow('m014a0501.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','w014A05');">Add Backorder Received</a></td>
    </tr>
</table>
</body>
</html>
<%
rsBackOrders.Close();
%>