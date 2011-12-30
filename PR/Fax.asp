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

	if (rsVendor.Fields.Item("intPhone_Type_1").Value == 5) Fax = "(" + rsVendor.Fields.Item("chvPhone1_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone1_Num").Value;
	if (rsVendor.Fields.Item("intPhone_Type_1").Value == 2) Phone = "(" + rsVendor.Fields.Item("chvPhone1_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone1_Num").Value;
	if (rsVendor.Fields.Item("intPhone_Type_2").Value == 5) Fax = "(" + rsVendor.Fields.Item("chvPhone2_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone2_Num").Value;
	if (rsVendor.Fields.Item("intPhone_Type_2").Value == 2) Phone = "(" + rsVendor.Fields.Item("chvPhone2_Arcd").Value + ") " + rsVendor.Fields.Item("chvPhone2_Num").Value;
		
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
%>
<%Response.ContentType = "application/msword"%>
<html>
<head>
	<title>Confidential Facsimile</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" style="font-face: arial">
<p align="center"><img src="http://<%=Request.ServerVariables("server_name")%>:8080/aspsite/i/letterhead.gif" width="450" height="80"></p>
<p align="center" style="font-size: 15pt"><b>CONFIDENTIAL FACSIMILE</b></p>
<table align="center" cellpadding="3" cellspacing="3" width="480">
	<tr>
		<td width="70">Date:</td>
		<td width="170"><%=CurrentDate()%></td>
		<td width="70"></td>
		<td width="170"></td>
	</tr>
	<tr>
		<td>Attention:</td>
		<td><%=ContactName%></td>
		<td>Fax</td>
		<td><%=Fax%></td>
	</tr>
	<tr>
		<td></td>
		<td></td>
		<td>Phone</td>
		<td><%=Phone%></td>
	</tr>
	<tr>
		<td>From:</td>
		<td><%=StaffName%></td>
		<td>Pages:</td>
		<td>2 (including this page)</td>
	</tr>
</table>
<br>
RE: Purchase Order #&nbsp;<%=PurchaseOrderNumber%><br>
<hr>
<br>
<br>
Please proceed with the attached purchase requisition # <%=Request.QueryString("insPurchase_Req_id")%> and ship to above address.<br>
<br>
If you have any questions, please give me a call at (604) 959-8188.<br>
<br>
Thank you.<br>
<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>
<table cellpadding="3" cellspacing="1" align="center" style="border: 1px solid">
	<tr>
		<td align="center"><i>If you do not receive all pages or copy is difficult to read,<br>please call for immediate re-transmission.</i></td>
	</tr>
</table>
</body>
</html>
