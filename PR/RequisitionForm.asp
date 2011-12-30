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
var DateOrdered = "";
var OrderedBy = "";
var Vendor = "";
var VendorAddress = "";

WorkOrderNumber = rsPurchaseRequisition.Fields.Item("chvWork_order").Value;
DateOrdered = FilterDate(rsPurchaseRequisition.Fields.Item("dtsDate_Ordered").Value);
OrderedBy = rsPurchaseRequisition.Fields.Item("chvOrdered_by").Value;
Vendor = rsPurchaseHeader.Fields.Item("chvVendor").Value;

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

var RequestNotes = "";
var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_Purchase_Requisition_Note(" + Request.QueryString("insPurchase_Req_id") + ",'',0,0,'',0,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();
while (!rsNotes.EOF) {
	if (rsNotes.Fields.Item("chvType_of_Note").Value == "Requested") RequestNotes = rsNotes.Fields.Item("chvNote_Desc").Value;
	rsNotes.MoveNext
}

var rsInventoryRequested = Server.CreateObject("ADODB.Recordset")
rsInventoryRequested.ActiveConnection = MM_cnnASP02_STRING
rsInventoryRequested.Source = "{call dbo.cp_Purchase_Requisition_Requested(0," + Request.QueryString("insPurchase_Req_id") + ",0,0,0,'',0.0,'01/01/1999',0,0,0,'Q',0)}"
rsInventoryRequested.CursorType = 0
rsInventoryRequested.CursorLocation = 2
rsInventoryRequested.LockType = 3
rsInventoryRequested.Open()
%>
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Purchase Requisition Form</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><img src="../i/letterhead.gif" width="339" height="63"></p>
<p align="center" style="font: 14pt Arial">PURCHASE REQUISITION FORM</p>
<table align="center" cellpadding="2" cellspacing="0" width="600" border="1" bordercolor="#CCCCCC" style="font: 10pt Arial">
	<tr>
		<td width="280">Work Order Number:&nbsp;<%=WorkOrderNumber%></td>
		<td>Purchase Order Number:&nbsp;<%=PurchaseOrderNumber%></td>
	</tr>
	<tr>
		<td valign="top">
			<table cellpadding="2" width="100%" cellspacing="1" style="font: 10pt Arial">
				<tr>
					<td width="90">PR Number:</td>
					<td><%=Request.QueryString("insPurchase_req_id")%></td>
				</tr>
				<tr>
					<td>Date Ordered:</td>
					<td><%=DateOrdered%></td>
				</tr>
				<tr>
					<td>Ordered By:</td>
					<td><%=OrderedBy%></td>
				</tr>
			</table>
		</td>
		<td valign="top">
			<table cellpadding="2" width="100%" cellspacing="1" style="font: 10pt Arial">
				<tr>
					<td valign="top" width="70">Vendor:</td>
					<td>
						<%=Vendor%><br>
						<%=VendorAddress%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<table align="center" cellpadding="2" cellspacing="0" width="600" border="1" bordercolor="#CCCCCC" style="font: 10pt Arial">
	<tr>
		<td colspan="4">Purchase Requested</td>
	</tr>
	<tr>
		<td align="center" width="50">Quantity</td>
		<td align="left" width="370">Description/Specification</td>
		<td align="center">List Unit Cost</td>
		<td align="center">Total Cost</td>
	</tr>
	<tr>
		<td height="260" colspan="4" valign="top">
			<table align="center" cellpadding="2" cellspacing="0" width="100%" border="0" style="font: 10pt Arial">
<%
while (!rsInventoryRequested.EOF) {
%>			
				<tr>
					<td width="55" valign="top" align="center"><%=rsInventoryRequested.Fields.Item("insPR_request_Qty_Ordered").Value%></td>
					<td width="365" valign="top" align="left"><%=rsInventoryRequested.Fields.Item("chvClass_Bundle_Name").Value%>&nbsp;<%=rsInventoryRequested.Fields.Item("chvDescription").Value%></td>
					<td width="100" valign="top" align="right"><%=FormatCurrency(rsInventoryRequested.Fields.Item("fltPR_request_List_Unit_Cost").Value)%></td>
					<td valign="top" align="right"><%=FormatCurrency(rsInventoryRequested.Fields.Item("fltTotal_Cost").Value)%></td>
				</tr>
<%
	rsInventoryRequested.MoveNext();
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
	<tr height="70">
		<td><%=RequestNotes%></td>
	</tr>
</table>
<br>
<table align="center" cellpadding="2" cellspacing="0" width="600" border="1" bordercolor="#CCCCCC" style="font: 10pt Arial">
	<tr height="60">
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td align="center">Signature of Purchaser</td>		
		<td align="center">Signature of Administrator</td>
	</tr>
</table>
</body>
</html>
