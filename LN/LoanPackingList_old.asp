<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var intShip_dtl_id = 0;
if (!rsLoan.EOF) {
	if (rsLoan.Fields.Item("intShip_dtl_id").Value != null) 
	    intShip_dtl_id = rsLoan.Fields.Item("intShip_dtl_id").Value;
} 

var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_Req_id")+",0,'',0,0,'','',0,'Q',0)}";
rsInventoryLoaned.CursorType = 0;
rsInventoryLoaned.CursorLocation = 2;
rsInventoryLoaned.LockType = 3;
rsInventoryLoaned.Open();

var rsInventoryBackOrdered = Server.CreateObject("ADODB.Recordset");
rsInventoryBackOrdered.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryBackOrdered.Source = "{call dbo.cp_eqp_requested(0,"+Request.QueryString("intLoan_Req_id")+",0,0,0,'',0.0,0,2,'Q',0)}";
rsInventoryBackOrdered.CursorType = 0;
rsInventoryBackOrdered.CursorLocation = 2;
rsInventoryBackOrdered.LockType = 3;
rsInventoryBackOrdered.Open();

var rsAccessories = Server.CreateObject("ADODB.Recordset");
rsAccessories.ActiveConnection = MM_cnnASP02_STRING;
rsAccessories.Source = "{call dbo.cp_get_loan_accessory2("+Request.QueryString("intLoan_Req_id")+",0,0)}";
rsAccessories.CursorType = 0;
rsAccessories.CursorLocation = 2;
rsAccessories.LockType = 3;
rsAccessories.Open();

var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
rsShippingMethod.Source = "{call dbo.cp_loan_ship_method("+intShip_dtl_id+",0,'',0,0,0,'',0,'','',0,0,'Q',0)}";
rsShippingMethod.CursorType = 0;
rsShippingMethod.CursorLocation = 2;
rsShippingMethod.LockType = 3;
rsShippingMethod.Open();

var rsShippingAddress = Server.CreateObject("ADODB.Recordset");
rsShippingAddress.ActiveConnection = MM_cnnASP02_STRING;
rsShippingAddress.Source = "{call dbo.cp_loan_ship_address("+intShip_dtl_id+",0,'','','','','',0,'','',0,'',0,'','','',0,'','','',0,'','','','','',0,'Q',0)}";
rsShippingAddress.CursorType = 0;
rsShippingAddress.CursorLocation = 2;
rsShippingAddress.LockType = 3;
rsShippingAddress.Open();
%>
<html>
<head>
	<title>Loan Packing List</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
    <style type="text/css">
<!--
.style1 {
	color: #0000FF;
	font-weight: bold;
}
-->
    </style>
</head>
<body>
<span class="style1">Equipment Loaned</span><br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><strong>Inventory Name</strong></td>
		<td><strong>Equipment ID</strong></td>
		<td><strong>Model Number</strong></td>
		<td><strong>Serial Number</strong></td>
		<td><strong>PR Number</strong></td>
		<td><strong>Equipment Cost</strong></td>										
	</tr>
<%
while (!rsInventoryLoaned.EOF) {
%>
	<tr>
		<td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvInventory_Name").Value)%></td>
		<td nowrap align="center"><%=ZeroPadFormat(rsInventoryLoaned.Fields.Item("intBar_Code_no").Value,8)%></td>
		<td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvModel_Number").Value)%></td>
		<td nowrap align="left"><%=(rsInventoryLoaned.Fields.Item("chvSerial_Number").Value)%></td>
		<td nowrap align="center"><%=ZeroPadFormat(rsInventoryLoaned.Fields.Item("intRequisition_no").Value,8)%></td>
		<td nowrap align="right"><%=FormatCurrency(rsInventoryLoaned.Fields.Item("fltList_Unit_Cost").Value)%></td>
	</tr>
<%
	rsInventoryLoaned.MoveNext();
}
%>
</table>
<br>
<span class="style1">Backorder List</span><br>
<table cellpadding="1" cellspacing="1" border="1" style="border: solid 1px #CCCCCC">
	<tr>
		<td><strong>Inventory/Bundle Name</strong></td>
	</tr>
<%
while (!rsInventoryBackOrdered.EOF) {
%>
	<tr>
		<td nowrap align="left"><%=((rsInventoryBackOrdered.Fields.Item("bitIs_class").Value)?rsInventoryBackOrdered.Fields.Item("chvEqp_Class_Name").Value:rsInventoryBackOrdered.Fields.Item("chvEqp_Bundle_Name").Value)%></td>
	</tr>
<%
	rsInventoryBackOrdered.MoveNext();
}
%>
</table>
<br>
<span class="style1">Accessories Included</span><br>
<table cellpadding="1" cellspacing="1" border="1" style="border: solid 1px #CCCCCC">
	<tr>
		<td><strong>Item</strong></td>
		<td><strong>Qty.</strong></td>		
	</tr>
<%
while (!rsAccessories.EOF) {
%>
	<tr>	
		<td><%=rsAccessories.Fields.Item("chvAttach_name").Value%></td>
		<td><%=rsAccessories.Fields.Item("insQuantity").Value%></td>		
	</tr>
<%
	rsAccessories.MoveNext();
}
%>
</table>
<br>
<b>Shipping Information</b><br>
Delivered:&nbsp;
<!-- + Nov.03.2005
-->
<%
var stgTmp
//FilterDate(rsShippingMethod.Fields.Item("dtsDlvy_date").Value)
if (!rsLoan.EOF) {
    stgTmp = FilterDate(rsShippingMethod.Fields.Item("dtsDlvy_date").Value) ;
    Response.Write(stgTmp)
}  
%>
&nbsp;<span style="font-size: 7pt">(mm/dd/yyyy)</span><br>

<!-- + Nov.03.2005
-->
Shipping Method:&nbsp;
<%
//Shipping Method:&nbsp;<%=rsShippingMethod.Fields.Item("chvShip_Method").Value%><br>
if (!rsLoan.EOF) {
   Response.Write(rsShippingMethod.Fields.Item("chvShip_Method").Value)
}  
%>
<br>

Waybill Number:&nbsp;
<%
//<%=rsShippingMethod.Fields.Item("chvWayBill_No").Value%>
if (!rsLoan.EOF) {
   Response.Write(rsShippingMethod.Fields.Item("chvWayBill_No").Value)
}
%>

<br>
<br>
<br>
<b>Shipping Address</b><br>

Name:&nbsp;
<%
// + Nov.03.2005
//<%=rsShippingAddress.Fields.Item("chvUsr_Fstname").Value%>
if (!rsShippingAddress.EOF) {
  Response.Write(rsShippingAddress.Fields.Item("chvUsr_Lstname").Value)
}
%>
&nbsp;<br>

C/O:&nbsp;
<%
// + Nov.03.2005
//<%=rsShippingAddress.Fields.Item("chvCo_Fstname").Value%>
if (!rsShippingAddress.EOF) {
  Response.Write(rsShippingAddress.Fields.Item("chvCo_Fstname").Value)
}
%>
&nbsp;
<%
// + Nov.03.2005
//<%=rsShippingAddress.Fields.Item("chvCo_Lstname").Value%>
if (!rsShippingAddress.EOF) {
  Response.Write(rsShippingAddress.Fields.Item("chvCo_Lstname").Value)
}
%>
<br>
Facility:&nbsp;
<%
// + Nov.03.2005
//<%=rsShippingAddress.Fields.Item("chvFacility").Value%>
if (!rsShippingAddress.EOF) {
  Response.Write(rsShippingAddress.Fields.Item("chvFacility").Value)
}
%>
<br>
Address:&nbsp;
<%
// + Nov.03.2005
//<%=rsShippingAddress.Fields.Item("chvAddress").Value%>
if (!rsShippingAddress.EOF) {
  Response.Write(rsShippingAddress.Fields.Item("chvAddress").Value)
}
%>
<br>
City:&nbsp;
<%
// + Nov.03.2005
//<%=rsShippingAddress.Fields.Item("chvCity").Value%>
if (!rsShippingAddress.EOF) {
  Response.Write(rsShippingAddress.Fields.Item("chvCity").Value)
}
%>
<br>
Postal Code:&nbsp;
<%
// + Nov.03.2005
//<%=FormatPostalCode(rsShippingAddress.Fields.Item("chvPostal_zip").Value)%>
if (!rsShippingAddress.EOF) {
  Response.Write(FormatPostalCode(rsShippingAddress.Fields.Item("chvPostal_zip").Value))
}
%>
<br>
Phone Number:&nbsp;
<%
// + Nov.03.2005
//<%=FormatPhoneNumber(rsShippingAddress.Fields.Item("chvPhone_Type_1").Value,rsShippingAddress.Fields.Item("chvPhone1_Arcd").Value,rsShippingAddress.Fields.Item("chvPhone1_Num").Value,rsShippingAddress.Fields.Item("chvPhone1_Ext").Value,rsShippingAddress.Fields.Item("chvPhone_Type_2").Value,rsShippingAddress.Fields.Item("chvPhone2_Arcd").Value,rsShippingAddress.Fields.Item("chvPhone2_Num").Value,rsShippingAddress.Fields.Item("chvPhone2_Ext").Value,"","","","","")%>
if (!rsShippingAddress.EOF) {
  Response.Write(FormatPhoneNumber(rsShippingAddress.Fields.Item("chvPhone_Type_1").Value,rsShippingAddress.Fields.Item("chvPhone1_Arcd").Value,rsShippingAddress.Fields.Item("chvPhone1_Num").Value,rsShippingAddress.Fields.Item("chvPhone1_Ext").Value,rsShippingAddress.Fields.Item("chvPhone_Type_2").Value,rsShippingAddress.Fields.Item("chvPhone2_Arcd").Value,rsShippingAddress.Fields.Item("chvPhone2_Num").Value,rsShippingAddress.Fields.Item("chvPhone2_Ext").Value,"","","","",""))
}
%>

<br>

</body>
</html>
<%
rsAccessories.Close();
rsInventoryBackOrdered.Close();
rsInventoryLoaned.Close();
%>