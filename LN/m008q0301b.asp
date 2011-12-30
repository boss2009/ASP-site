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

var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryLoaned.Source = "select * from tbl_equip_loaned where intloan_req_id = " + Request.QueryString("intLoan_req_id");
rsInventoryLoaned.CursorType = 0;
rsInventoryLoaned.CursorLocation = 2;
rsInventoryLoaned.LockType = 3;
rsInventoryLoaned.Open();

var rsInventoryLoaned_total = 0;
while (!rsInventoryLoaned.EOF) {
	rsInventoryLoaned_total++;
	rsInventoryLoaned.MoveNext();
}
rsInventoryLoaned.Requery();
%>
<html>
<head>
	<title>Equipment Loaned</title>
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
<h5>Equipment Loaned</h5>
<table cellspacing="1">
    <tr> 
		<td nowrap width="450"><a href="javascript: openWindow('m008a0301.asp?intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>','w008A02');">Add Equipment Loaned</a></td>	
    	<td nowrap align="left">Displaying <b><%=(rsInventoryLoaned_total)%></b> Records.</td>
	</tr>
</table>
<hr>
<div class="BrowsePanel" style="width: 564px; height: 219px"> 
  <table cellpadding="2" cellspacing="1">
    <tr> 
      <th nowrap class="headrow" align="left">Inventory Name</th>
      <th nowrap class="headrow" align="left">Inventory ID</th>
      <th nowrap class="headrow" align="left">Date Returned</th>
      <th nowrap class="headrow" align="left">Status</th>
      <th nowrap class="headrow" align="left">Vendor</th>
      <th nowrap class="headrow" align="left">Model Number</th>
      <th nowrap class="headrow" align="left">Serial Number</th>
      <th nowrap class="headrow" align="left">PR Number</th>
      <th nowrap class="headrow" align="left">Equipment Cost</th>
      <th nowrap class="headrow" align="left">Date Processed</th>
      <th nowrap class="headrow" align="left">Returned By</th>
      <th nowrap class="headrow" align="left">Return Status</th>
      <th nowrap class="headrow" align="left">Comments</th>
    </tr>
<% 
var tax = 0;
var total_shipping = 0;
var total_cost = 0;
var total_loan = 0;
var Vendor = "";
var ReturnedBy = "";
while ((!rsInventoryLoaned.EOF)) {
	var rsInventoryDetail = Server.CreateObject("ADODB.Recordset");
	rsInventoryDetail.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryDetail.Source = "{call dbo.cp_Get_EqCls_Inventory(1,0,'',1," + rsInventoryLoaned.Fields.Item("intEquip_set_id").Value + ",0)}";	
	rsInventoryDetail.CursorType = 0;
	rsInventoryDetail.CursorLocation = 2;
	rsInventoryDetail.LockType = 3;
	rsInventoryDetail.Open();	

	if ((rsInventoryDetail.Fields.Item("intRequisition_no").Value > 0) && (rsInventoryDetail.Fields.Item("intRequisition_no").Value < 30000)) {	 
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
			
	if (rsInventoryLoaned.Fields.Item("insReturned_by_id").Value > 0) {
		var rsStaff = Server.CreateObject("ADODB.Recordset");
		rsStaff.ActiveConnection = MM_cnnASP02_STRING;
		rsStaff.Source = "{call dbo.cp_staff2("+rsInventoryLoaned.Fields.Item("insReturned_by_id").Value+",0,'','',0,'','',0,0,0,0,0,0,0,0,0,1,0,'',1,'Q',0)}"
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
			tax = tax + (rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value * (pst/100));
		break;
		//gst
		case "2":
			tax = tax + (rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value * (gst/100));		
		break;
		//both
		case "3":
			tax = tax + (rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value * ((gst+pst)/100));		
		break;
	}
%>
    <tr> 
      <td nowrap align="left"><a href="m008e0301.asp?intEqp_Loaned_Id=<%=rsInventoryLoaned.Fields.Item("intEqp_Loaned_Id").Value%>&intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>"><%=(rsInventoryDetail.Fields.Item("chvInventory_Name").Value)%></a>&nbsp;</td>
      <td nowrap align="center"><%=ZeroPadFormat(rsInventoryDetail.Fields.Item("intBar_Code_no").Value,8)%>&nbsp;</td>
      <td nowrap align="center"><%=FilterDate(rsInventoryLoaned.Fields.Item("dtsDate_Returned").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryDetail.Fields.Item("chvEqp_Status").Value)%>&nbsp;</td>	  	  
      <td nowrap align="left"><%=(Vendor)%>&nbsp;</td>	  	  
      <td nowrap align="left"><%=(rsInventoryDetail.Fields.Item("chvModel_Number").Value)%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryDetail.Fields.Item("chvSerial_Number").Value)%>&nbsp;</td>	  
      <td nowrap align="center"><%=ZeroPadFormat(rsInventoryDetail.Fields.Item("intRequisition_no").Value,8)%>&nbsp;</td>
      <td nowrap align="right"><%=FormatCurrency(rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value)%>&nbsp;</td>	  
      <td nowrap align="center"><%=FilterDate(rsInventoryLoaned.Fields.Item("dtsDate_Shipped").Value)%>&nbsp;</td>	  
      <td nowrap align="left"><%=(ReturnedBy)%>&nbsp;</td>	  
      <td nowrap align="center"><%=((rsInventoryLoaned.Fields.Item("bitRtn_Complete").Value=="1")?"Complete":"Incomplete")%>&nbsp;</td>
      <td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvComments").Value)%>&nbsp;</td>
    </tr>
<%
if (rsInventoryLoaned.Fields.Item("dtsDate_Returned").Value==null) total_loan += rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value;
	total_cost += rsInventoryDetail.Fields.Item("fltPurchase_Cost").Value;
	rsInventoryLoaned.MoveNext();
}

total_shipping = total_cost * (shipping/100);
%>
  </table>
</div>
<div style="position: absolute; top: 310px">
<table cellpadding="0" cellspacing="1">
	<tr>
		<td width="350"><b>Total Loan Cost with taxes/shipping:</b></td>
		<td align="right"><b><%=FormatCurrency(total_cost+ tax + total_shipping)%></b></td>
	</tr>
	<tr>
		<td width="350"><b>Total Cost of Equipment Still On Loan:</b></td>
		<td align="right"><b><%=FormatCurrency(total_loan)%></b></td>
	</tr>
</table>
</div>
</body>
</html>
<%
rsInventoryLoaned.Close();
%>