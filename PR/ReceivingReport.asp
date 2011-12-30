<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var rsPurchaseRequisition = Server.CreateObject("ADODB.Recordset");
rsPurchaseRequisition.ActiveConnection = MM_cnnASP02_STRING;
rsPurchaseRequisition.Source = "{call dbo.cp_Get_Purchase_Requisition_02(0,0,'',1," + Request("insPurchase_Req_id") + ",0)}";
rsPurchaseRequisition.CursorType = 0;
rsPurchaseRequisition.CursorLocation = 2;
rsPurchaseRequisition.LockType = 3;
rsPurchaseRequisition.Open();

var rsPurchaseHeader = Server.CreateObject("ADODB.Recordset");
rsPurchaseHeader.ActiveConnection = MM_cnnASP02_STRING;
rsPurchaseHeader.Source = "{call dbo.cp_FrmHdr(14,"+Request.QueryString("insPurchase_Req_id")+")}";
rsPurchaseHeader.CursorType = 0;
rsPurchaseHeader.CursorLocation = 2;
rsPurchaseHeader.LockType = 3;
rsPurchaseHeader.Open();

var WorkOrderNumber = "";
var DateReceived = "";
var OrderedBy = "";
var Vendor = "";

WorkOrderNumber = rsPurchaseRequisition.Fields.Item("chvWork_order").Value;
DateReceived = FilterDate(rsPurchaseRequisition.Fields.Item("dtsDate_Received").Value);
Vendor = rsPurchaseRequisition.Fields.Item("chvSupplier").Value;

var rsContractPO = Server.CreateObject("ADODB.Recordset")
rsContractPO.ActiveConnection = MM_cnnASP02_STRING
rsContractPO.Source = "{call dbo.cp_Purchase_Requisition_Vendor(" + Request.QueryString("insPurchase_Req_id") + ",0)}"
rsContractPO.CursorType = 0
rsContractPO.CursorLocation = 2
rsContractPO.LockType = 3
rsContractPO.Open()

var PurchaseOrderNumber = "";

if (rsPurchaseRequisition.Fields.Item("insRequest_type_id").Value>="6") {
	PurchaseOrderNumber	= "Purchase Card";
} else {
	if (!rsContractPO.EOF) PurchaseOrderNumber = rsContractPO.Fields.Item("chvContract_PO").Value;
}


var ContactName = ""
var Fax = ""
var Phone = ""
if (rsPurchaseHeader.Fields.Item("insVendor_id").Value > 0) {
	var rsVendor = Server.CreateObject("ADODB.Recordset")
	rsVendor.ActiveConnection = MM_cnnASP02_STRING
	rsVendor.Source = "{call dbo.cp_Get_Company_Address(" + rsPurchaseHeader.Fields.Item("insVendor_id").Value + ", 1)}"
	rsVendor.CursorType = 0
	rsVendor.CursorLocation = 2
	rsVendor.LockType = 3
	rsVendor.Open()

	VendorAddress = rsVendor.Fields.Item("chvAddress").Value + "<br>" + rsVendor.Fields.Item("chvCity").Value + ", " + rsVendor.Fields.Item("chrprvst_abbv").Value + "<br>";
	VendorAddress = VendorAddress + FormatPostalCode(rsVendor.Fields.Item("chvPostal_zip").Value) + "<br>"

	if (rsVendor.Fields.Item("intPhone_Type_1").Value == 5) Fax = "(" + rsVendor.Fields.Item("chvPhone1_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone1_Num").Value;
	if (rsVendor.Fields.Item("intPhone_Type_1").Value == 2) Phone = "(" + rsVendor.Fields.Item("chvPhone1_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone1_Num").Value;
	if (rsVendor.Fields.Item("intPhone_Type_2").Value == 5) Fax = "(" + rsVendor.Fields.Item("chvPhone2_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone2_Num").Value;
	if (rsVendor.Fields.Item("intPhone_Type_2").Value == 2) Phone = "(" + rsVendor.Fields.Item("chvPhone2_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone2_Num").Value;

	if (Phone != "") VendorAddress = VendorAddress + "Off: " + Phone + "<br>";
	if (Fax != "") VendorAddress = VendorAddress + "FAX: " + Fax
		
	var rsVendorContact = Server.CreateObject("ADODB.Recordset");
	rsVendorContact.ActiveConnection = MM_cnnASP02_STRING;
	rsVendorContact.Source = "{call dbo.cp_Get_Company_Address_KeyContact(" + rsPurchaseHeader.Fields.Item("insVendor_id").Value + ", 1, 1)}";
	rsVendorContact.CursorType = 0;
	rsVendorContact.CursorLocation = 2;
	rsVendorContact.LockType = 3;
	rsVendorContact.Open();
	
	if (!rsVendorContact.EOF) ContactName = rsVendorContact.Fields.Item("chvkeyContact_Fname").Value + " " + rsVendorContact.Fields.Item("chvkeyContact_Lname").Value;
}

var StaffName = "";
var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_Idv_Staff(" + Session("insStaff_id") + ")}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
StaffName = rsStaff.Fields.Item("chvFst_Name").Value + " " + rsStaff.Fields.Item("chvLst_Name").Value;

var ReceiveNotes = "";
var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_Purchase_Requisition_Note(" + Request.QueryString("insPurchase_Req_id") + ",'',0,0,'',0,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();
while (!rsNotes.EOF) {
	if (rsNotes.Fields.Item("chvType_of_Note").Value != "Requested") ReceiveNotes = rsNotes.Fields.Item("chvNote_Desc").Value;
	rsNotes.MoveNext
}

var rsInventoryReceived = Server.CreateObject("ADODB.Recordset");
rsInventoryReceived.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryReceived.Source = "{call dbo.cp_Purchase_Requisition_Received(" + Request("insPurchase_Req_id") + ",0,0,'',0,'',0,'Q',0)}";
rsInventoryReceived.CursorType = 0;
rsInventoryReceived.CursorLocation = 2;
rsInventoryReceived.LockType = 3;
rsInventoryReceived.Open();
%>
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Purchase Receiving Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><img src="http://<%=Request.ServerVariables("server_name")%>:8080/aspsite/i/letterhead.gif" width="450" height="80"></p>
<p align="center" style="font: 14pt Arial">PURCHASE RECEIVING REPORT</p>
<table align="center" cellpadding="2" cellspacing="0" width="500" style="font: 10pt Arial">
	<tr>
		<td nowrap>Ordered From:</td>
		<td nowrap><u><%=Vendor%></u></td>
		<td nowrap>Date Received:</td>
		<td nowrap><u><%=DateReceived%></u></td>
	</tr>
	<tr>
		<td nowrap>Purchase Order No.:</td>
		<td nowrap><u><%=PurchaseOrderNumber%></u></td>
		<td nowrap>PR Number:</td>
		<td nowrap><u><%=Request.QueryString("insPurchase_Req_id")%></u></td>
	</tr>
</table>
<br>
<table align="center" cellpadding="2" cellspacing="0" width="600" border="1" bordercolor="#CCCCCC" style="font: 10pt Arial">
	<tr>
		<td align="center" width="50">Qty. Received</td>
<!--	<td align="center" width="50">Qty. B/O</td>-->
		<td align="left" width="300">Description</td>
		<td align="left">Remarks</td>
	</tr>
	<tr>
		<td height="420" colspan="4" valign="top">
			<table align="center" cellpadding="2" cellspacing="0" width="100%" border="0" style="font: 10pt Arial">
<%
while (!rsInventoryReceived.EOF) {
%>			
				<tr>
					<td width="55" valign="top" align="center"><%=rsInventoryReceived.Fields.Item("intQuantity_Ordered").Value%></td>
<!--				<td width="50" valign="top" align="center"><%=rsInventoryReceived.Fields.Item("intBack_Order").Value%></td>-->
					<td width="300" valign="top" align="left"><%=rsInventoryReceived.Fields.Item("chvClass_name").Value%></td>
					<td valign="top" align="left"><%=rsInventoryReceived.Fields.Item("chvNotes").Value%></td>
				</tr>
<%
	rsInventoryReceived.MoveNext();
}
%>			
			</table>	
		</td>
	</tr>
</table>
<br>
<table align="center" cellpadding="2" cellspacing="0" width="600" border="1" bordercolor="#CCCCCC" style="font: 10pt Arial">
	<tr>
		<td>Notes</td>
	</tr>
	<tr height="90">
		<td><%=ReceiveNotes%></td>
	</tr>
</table>
<br>
<br>
&nbsp;&nbsp;&nbsp;&nbsp;Goods received as checked above ____________________________
</body>
</html>