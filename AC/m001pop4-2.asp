<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE file="../Connections/cnnASP02.asp" -->
<!--#INCLUDE file="../inc/ASPCheckLogin.inc" -->
<%
var gst = 0;
var pst = 0;
var shipping = 0;

var rsGST = Server.CreateObject("ADODB.Recordset");
rsGST.ActiveConnection = MM_cnnASP02_STRING;
rsGST.Source = "{call dbo.cp_charge_rate(1,'',0,0.0,1,'Q',0)}";
rsGST.CursorType = 0;
rsGST.CursorLocation = 2;
rsGST.LockType = 3;
rsGST.Open();
if (!rsGST.EOF) gst = Number(rsGST.Fields.Item("fltPercentage").Value);
rsGST.Close();

var rsPST = Server.CreateObject("ADODB.Recordset");
rsPST.ActiveConnection = MM_cnnASP02_STRING;
rsPST.Source = "{call dbo.cp_charge_rate(2,'',0,0.0,1,'Q',0)}";
rsPST.CursorType = 0;
rsPST.CursorLocation = 2;
rsPST.LockType = 3;
rsPST.Open();
if (!rsPST.EOF) pst = Number(rsPST.Fields.Item("fltPercentage").Value);
rsPST.Close();

var rsShipping = Server.CreateObject("ADODB.Recordset");
rsShipping.ActiveConnection = MM_cnnASP02_STRING;
rsShipping.Source = "{call dbo.cp_charge_rate(3,'',0,0.0,1,'Q',0)}";
rsShipping.CursorType = 0;
rsShipping.CursorLocation = 2;
rsShipping.LockType = 3;
rsShipping.Open();
if (!rsShipping.EOF) shipping = Number(rsShipping.Fields.Item("fltPercentage").Value);
rsShipping.Close();

var rsBuyoutSummary = Server.CreateObject("ADODB.Recordset");
rsBuyoutSummary.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutSummary.Source = "{call dbo.cp_buyout_hstry_summary("+ Request.QueryString("intAdult_id") + ",1,1,1,0)}";
rsBuyoutSummary.CursorType = 0;
rsBuyoutSummary.CursorLocation = 2;
rsBuyoutSummary.LockType = 3;
rsBuyoutSummary.Open();
%>
<html>
<head>
	<title>Buyout Summary</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<h5>Buyout Summary</h5>
<hr>
<% 
while (!rsBuyoutSummary.EOF) { 
	var rsInventorySold = Server.CreateObject("ADODB.Recordset");
	rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
	rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold(0,"+rsBuyoutSummary.Fields.Item("intBuyout_req_id").Value+",0,0.0,'',0,0,'',0,'Q',0)}";
	rsInventorySold.CursorType = 0;
	rsInventorySold.CursorLocation = 2;
	rsInventorySold.LockType = 3;
	rsInventorySold.Open();
%>
<b>Buyout ID: <%=(rsBuyoutSummary.Fields.Item("intBuyout_req_id").Value)%></b>
<table cellspacing="1" cellpadding="2" class="Mtable">
    <tr> 
		<th class="headrow" valign="top" align="left" width="300">Inventory Name</th>
		<th class="headrow" valign="top" align="left">Inventory ID</th>								
		<th class="headrow" valign="top" align="left">Date Processed</th>
		<th class="headrow" valign="top" align="left">Date Delivered</th>
		<th class="headrow" valign="top" align="left">Date Returned</th>						
    </tr>
<%
var total_sold_price = 0;
var total_cost = 0;
var tax = 0;
var total_shipping = 0;
while (!rsInventorySold.EOF) { 
	if (!(rsInventorySold.Fields.Item("insEquip_Class_id").Value==null)) {		
		var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
		rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
		rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW(" + rsInventorySold.Fields.Item("insEquip_Class_id").Value + ",'C',1)}";	
		rsConcreteClass.CursorType = 0;
		rsConcreteClass.CursorLocation = 2;
		rsConcreteClass.LockType = 3;
		rsConcreteClass.Open();	
		switch (String(rsConcreteClass.Fields.Item("chvSbjTotax").Value)) {
			//pst
			case "1":
				tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * (pst/100));
			break;
			//gst
			case "2":
				tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * (gst/100));		
			break;
			//both
			case "3":
				tax = tax + (rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value * ((gst+pst)/100));		
			break;
		}
	}
	total_cost += rsInventorySold.Fields.Item("fltList_Unit_Cost").Value;
	total_sold_price += rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value;		
%>
    <tr> 
		<td valign="top"><%=(rsInventorySold.Fields.Item("chvInventory_Name").Value)%>&nbsp;</td>
		<td nowrap valign="top" align="center"><%=ZeroPadFormat(rsInventorySold.Fields.Item("intEquip_Set_id").Value,8)%></td>						
		<td nowrap valign="top" align="center"><%=FilterDate(rsInventorySold.Fields.Item("dtsDate_processed").Value)%></td>
		<td nowrap valign="top" align="center"><%=FilterDate(rsInventorySold.Fields.Item("dtsDlvy_date").Value)%></td>		
		<td nowrap valign="top" align="center"><%=FilterDate(rsInventorySold.Fields.Item("dtsDate_Returned").Value)%></td>	
    </tr>
<%
	rsInventorySold.MoveNext();
}
total_shipping = total_sold_price * (shipping/100);
%>		
</table><br>
<b>
Total Sold Cost: <%=FormatCurrency(total_sold_price)%><br>
Taxes: <%=FormatCurrency(tax)%><br>
Shipping: <%=FormatCurrency(total_shipping)%><br>
Grand Total Sold Cost: <%=FormatCurrency(total_sold_price+tax+total_shipping)%><br>
</b><br>
<%
	rsBuyoutSummary.MoveNext();
}
%>
<hr>
<input type="button" value="Close" onClick="window.close();" class="btnstyle">
</body>
</html>
<%
rsBuyoutSummary.Close();
%>