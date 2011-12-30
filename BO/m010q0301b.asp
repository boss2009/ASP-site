<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
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

var rsInventorySold = Server.CreateObject("ADODB.Recordset");
rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
rsInventorySold.Source = "select * from tbl_buyout_equip_sold where intBuyout_req_id = " + Request.QueryString("intBuyout_req_id");
rsInventorySold.CursorType = 0;
rsInventorySold.CursorLocation = 2;
rsInventorySold.LockType = 3;
rsInventorySold.Open();
var rsInventorySold_total = 0;
while (!rsInventorySold.EOF) {
	rsInventorySold_total++;
	rsInventorySold.MoveNext();
}
rsInventorySold.Requery();
%>
<html>
<head>
	<title>Equipment Sold</title>
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
<h5>Equipment Sold</h5>
<table cellspacing="1">
    <tr> 
		<td nowrap width="450"><a href="javascript: openWindow('m010a0301.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>','w010A02');">Add Equipment Sold</a></td>			
    	<td nowrap align="left">Displaying <b><%=(rsInventorySold_total)%></b> Records.</td>
	</tr>
</table>
<hr>
<div class="BrowsePanel" style="width: 100%; height: 221px"> 
  <table cellpadding="2" cellspacing="1">
    <tr> 
		<th nowrap class="headrow" align="left">Inventory Name</th>
		<th nowrap class="headrow" align="left">Inventory ID</th>
		<th nowrap class="headrow" align="left">Status</th>
		<th nowrap class="headrow" align="left">Date Processed</th>
		<th nowrap class="headrow" align="left">Sold Price</th>
		<th nowrap class="headrow" align="left">Equipment Cost</th>
		<th nowrap class="headrow" align="left">Serial Number</th>
		<th nowrap class="headrow" align="left">Model Number</th>		
		<th nowrap class="headrow" align="left">PR Number</th>
		<th nowrap class="headrow" align="left">Vendor</th>
		<th nowrap class="headrow" align="left">Date Returned</th>
		<th nowrap class="headrow" align="left">Returned By</th>
		<th nowrap class="headrow" align="left">Comments</th>
    </tr>
<% 
//Getting Date Processed from shipping detail
var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var intShip_dtl_id = 0;
if (!rsBuyout.EOF) {
	if (rsBuyout.Fields.Item("intShip_dtl_id").Value != null) intShip_dtl_id = rsBuyout.Fields.Item("intShip_dtl_id").Value;
} 

var rsMethod = Server.CreateObject("ADODB.Recordset");
rsMethod.ActiveConnection = MM_cnnASP02_STRING;
rsMethod.Source = "{call dbo.cp_buyout_ship_method3("+ intShip_dtl_id + ",0,0,'',0,0,0,'',0,0,'','',0,0,'Q',0)}";
rsMethod.CursorType = 0;
rsMethod.CursorLocation = 2;
rsMethod.LockType = 3;
rsMethod.Open();

var IsNew = ((rsMethod.EOF)?true:false);

var total_sold_price = 0;
var total_cost = 0;
var tax = 0;
var total_shipping = 0;
var Vendor = "";
var ReturnedBy = "";
var DateProcessed = "";
if (!IsNew) DateProcessed = FilterDate(rsMethod.Fields.Item("dtsUser_Ship_date").Value);

while (!rsInventorySold.EOF) { 
	var rsInventoryDetail = Server.CreateObject("ADODB.Recordset");
	rsInventoryDetail.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryDetail.Source = "{call dbo.cp_Get_EqCls_Inventory(1,0,'',1," + rsInventorySold.Fields.Item("intEquip_set_id").Value + ",0)}";	
	rsInventoryDetail.CursorType = 0;
	rsInventoryDetail.CursorLocation = 2;
	rsInventoryDetail.LockType = 3;
	rsInventoryDetail.Open();

	if ((rsInventoryDetail.Fields.Item("intRequisition_no").Value > 0) && (rsInventoryDetail.Fields.Item("intRequisition_no").Value < 30000)){	 
		var rsPurchaseHeader = Server.CreateObject("ADODB.Recordset");
		rsPurchaseHeader.ActiveConnection = MM_cnnASP02_STRING;
		rsPurchaseHeader.Source = "{call dbo.cp_FrmHdr(14,"+rsInventoryDetail.Fields.Item("intRequisition_no").Value+")}";
		rsPurchaseHeader.CursorType = 0;
		rsPurchaseHeader.CursorLocation = 2;
		rsPurchaseHeader.LockType = 3;
		rsPurchaseHeader.Open();
		if (!rsPurchaseHeader.EOF) {
			Vendor = Trim(rsPurchaseHeader.Fields.Item("chvVendor").Value);
		} else {
			Vendor = "";
		}
	} else {
		Vendor = "";
	}
			
	if (rsInventorySold.Fields.Item("insRtned_by_id").Value > 0) {
		var rsStaff = Server.CreateObject("ADODB.Recordset");
		rsStaff.ActiveConnection = MM_cnnASP02_STRING;
		rsStaff.Source = "{call dbo.cp_staff2("+rsInventorySold.Fields.Item("insRtned_by_id").Value+",0,'','',0,'','',0,0,0,0,0,0,0,0,0,1,0,'',1,'Q',0)}"
		rsStaff.CursorType = 0;
		rsStaff.CursorLocation = 2;
		rsStaff.LockType = 3;
		rsStaff.Open();	
		if (!rsStaff.EOF) {
			ReturnedBy = Trim(rsStaff.Fields.Item("chvFst_Name").Value) + " " + Trim(rsStaff.Fields.Item("chvLst_Name").Value);
		} else {
			ReturnedBy = "";
		}
	} else {
		ReturnedBy = "";
	}
	
	switch (String(rsInventoryDetail.Fields.Item("chvSbjTotax").Value)) {
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

	total_cost += rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value;
	total_sold_price += rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value;		
%>
    <tr> 
		<td nowrap align="left"><a href="m010e0301.asp?intBO_Eqp_Sold_id=<%=rsInventorySold.Fields.Item("intBO_Eqp_Sold_id").Value%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>"><%=(rsInventoryDetail.Fields.Item("chvInventory_Name").Value)%></a>&nbsp;</td>
		<td nowrap align="center"><%=ZeroPadFormat(rsInventoryDetail.Fields.Item("intBar_Code_no").Value,8)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventoryDetail.Fields.Item("chvEqp_Status").Value)%>&nbsp;</td>		
		<td nowrap align="center"><%=DateProcessed%>&nbsp;</td>		
		<td nowrap align="right"><%=FormatCurrency(rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value)%>&nbsp;</td>				
		<td nowrap align="right"><%=FormatCurrency(rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value)%>&nbsp;</td>		
		<td nowrap align="left"><%=(rsInventoryDetail.Fields.Item("chvSerial_Number").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventoryDetail.Fields.Item("chvModel_Number").Value)%>&nbsp;</td>
		<td nowrap align="center"><%=ZeroPadFormat(rsInventoryDetail.Fields.Item("intRequisition_no").Value,8)%>&nbsp;</td>						
		<td nowrap align="left"><%=(Vendor)%>&nbsp;</td>		
		<td nowrap align="center"><%=FilterDate(rsInventorySold.Fields.Item("dtsDate_Returned").Value)%>&nbsp;</td>
		<td nowrap align="left"><%=(ReturnedBy)%>&nbsp;</td>
		<td nowrap align="left"><%=(rsInventorySold.Fields.Item("chvComments").Value)%>&nbsp;</td>
    </tr>
<%
	rsInventorySold.MoveNext();
}
total_shipping = total_sold_price * (shipping/100);
%>
  </table>
</div>
<div style="position: absolute; top: 310px">
<table cellpadding="0" cellspacing="1">
	<tr>
		<td width="350"><b>Total Equipment Cost:</b></td>
		<td align="right"><b><%=FormatCurrency(total_cost)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Total Sold Cost without taxes/shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_sold_price)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Taxes:</b></td>
		<td align="right"><b><%=FormatCurrency(tax)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_shipping)%></b></td>
	</tr>	
	<tr>	
		<td width="350"><b>Total Buyout Cost with taxes/shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_sold_price+tax+total_shipping)%></b></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsInventorySold.Close();
%>