<!--------------------------------------------------------------------------
* File Name: m014q01excel.asp
* Title: Purchase Requisition Browse
* Main SP: cp_get_purchase_requisition
* Description: This page lists purchase requisitions resulted from a search
* in excel format.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<% Response.ContentType = "application/vnd.ms-excel" %>
<%
var rsPurchase__inspSrtBy = "1";
if(String(Request.QueryString("inspSrtBy")) != "undefined") { 
  rsPurchase__inspSrtBy = String(Request.QueryString("inspSrtBy"));
}
var rsPurchase__inspSrtOrd = "0";
if(String(Request.QueryString("inspSrtOrd")) != "undefined") { 
  rsPurchase__inspSrtOrd = String(Request.QueryString("inspSrtOrd"));
}

var rsPurchase__chvFilter = "";
if(String(Request.QueryString("chvFilter")) != "undefined") { 
  rsPurchase__chvFilter = String(Request.QueryString("chvFilter"));
}

var rsPurchase = Server.CreateObject("ADODB.Recordset");
rsPurchase.ActiveConnection = MM_cnnASP02_STRING;
rsPurchase.Source = "{call dbo.cp_Get_Purchase_Requisition("+rsPurchase__inspSrtBy+","+rsPurchase__inspSrtOrd+",'"+rsPurchase__chvFilter.replace(/'/g, "''")+"',0,0,0)}";
rsPurchase.CursorType = 0;
rsPurchase.CursorLocation = 2;
rsPurchase.LockType = 3;
rsPurchase.Open();
%>
<html>
<head>
	<title>Purchase Requisition Browse</title>
</head>
<body>
<table>
	<tr> 
      <th>PR Number</th>
      <th>Vendor Name</th>
	  <th>Request Type</th>
      <th>Purchase Status</th>
      <th>Ordered Date</th>
      <th>Received Date</th>
      <th>On Backorder</th>
    </tr>
<% 
while (!rsPurchase.EOF) { 
%>
    <tr> 
      <td nowrap><%=ZeroPadFormat(rsPurchase.Fields.Item("insPurchase_Req_id").Value, 8)%></td>
      <td nowrap><%=(rsPurchase.Fields.Item("chvVendor").Value)%>&nbsp;</td>
      <td nowrap><%=(rsPurchase.Fields.Item("chvRequest_Type").Value)%>&nbsp;</td>
      <td nowrap><%=(rsPurchase.Fields.Item("chvPurchase_Status").Value)%>&nbsp;</td>
      <td nowrap><%=(rsPurchase.Fields.Item("dtsDate_Ordered").Value)%>&nbsp;</td>
      <td nowrap><%=(rsPurchase.Fields.Item("dtsDate_Received").Value)%>&nbsp;</td>
	  <td nowrap><%=(rsPurchase.Fields.Item("bitInv_on_bk_order").Value)%>&nbsp;</td>
    </tr>
<%
	rsPurchase.MoveNext();
}
%>
</table>
</body>
</html>
<%
rsPurchase.Close();
%>