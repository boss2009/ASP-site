<!--------------------------------------------------------------------------
* File Name: m014a0201.asp
* Title: New Inventory Request
* Main SP: cp_PR_request_validate, cp_purchase_requisition_requested
* Description: This page validates the equipment class first then inserts a 
* new inventory request.
* Author: T.H
--------------------------------------------------------------------------->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_Insert")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");			
	var ClassID = ((String(Request.Form("ClassID"))!="undefined")?0:Request.Form("ClassID"));	
	var Quantity = ((String(Request.Form("Quantity"))=="")?0:Request.Form("Quantity"));
	var ListUnitCost = ((String(Request.Form("ListUnitCost"))=="")?0:Request.Form("ListUnitCost"));
	var EstimatedDeliveryDate = ((String(Request.Form("EstimatedDeliveryDate"))!="undefined")?"":Request.Form("EstimatedDeliveryDate"));
				
	var ChkEquipClass = Server.CreateObject("ADODB.Command");
	ChkEquipClass.ActiveConnection = MM_cnnASP02_STRING;
	ChkEquipClass.CommandText = "dbo.cp_PR_request_Validate";
	ChkEquipClass.CommandType = 4;
	ChkEquipClass.CommandTimeout = 0;
	ChkEquipClass.Prepared = true;
	ChkEquipClass.Parameters.Append(ChkEquipClass.CreateParameter("RETURN_VALUE", 3, 4));
	ChkEquipClass.Parameters.Append(ChkEquipClass.CreateParameter("@insPurchase_Req_id", 3, 1,10000,Request.QueryString("insPurchase_Req_id")));
	ChkEquipClass.Parameters.Append(ChkEquipClass.CreateParameter("@insEquip_class_id", 3, 1,10000,Request.Form("ClassID")));	
	ChkEquipClass.Parameters.Append(ChkEquipClass.CreateParameter("@insRtnFlag", 2, 2));
	ChkEquipClass.Execute();	
	
	if (ChkEquipClass.Parameters.Item("@insRtnFlag").Value < 1) {
		var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
		rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
		//client
		if (String(Request.Form("UserType")) == "0") {
			rsInventoryRequest.Source = "{call dbo.cp_Purchase_Requisition_Requested2(0,"+Request.QueryString("insPurchase_Req_id")+","+Request.Form("ClassID")+",1,"+Request.Form("Quantity")+",'"+Description+"',"+Request.Form("ListUnitCost")+",'"+Request.Form("EstimatedDeliveryDate")+"',"+Request.Form("UserID")+",0,0,"+Request.Form("Vendor")+",0,'A',0)}";
		} else {
			rsInventoryRequest.Source = "{call dbo.cp_Purchase_Requisition_Requested2(0,"+Request.QueryString("insPurchase_Req_id")+","+Request.Form("ClassID")+",1,"+Request.Form("Quantity")+",'"+Description+"',"+Request.Form("ListUnitCost")+",'"+Request.Form("EstimatedDeliveryDate")+"',0,"+Request.Form("UserID")+",1,"+Request.Form("Vendor")+",0,'A',0)}";
		}
		rsInventoryRequest.CursorType = 0;
		rsInventoryRequest.CursorLocation = 2;
		rsInventoryRequest.LockType = 3;
	//	Response.Redirect(rsInventoryRequest.Source);
		rsInventoryRequest.Open();
		Response.Redirect("AddDeleteSuccessful.asp?action=Add");
	} else {
		Response.Redirect("AddDeleteFailed.asp?action=added");
	}
}

var ClassID = ((String(Request.QueryString("ClassID"))=="undefined")?0:Request.QueryString("ClassID"));

rsInventorySupplier = Server.CreateObject("ADODB.Recordset");
rsInventorySupplier.ActiveConnection = MM_cnnASP02_STRING;
rsInventorySupplier.Source = "{call dbo.cp_Get_EqCls_Vendor("+ClassID+",0,0)}";
rsInventorySupplier.CursorType = 0;
rsInventorySupplier.CursorLocation = 2;
rsInventorySupplier.LockType = 3;
rsInventorySupplier.Open();

var rsInventoryRequested = Server.CreateObject("ADODB.Recordset");
rsInventoryRequested.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryRequested.Source = "{call dbo.cp_Purchase_Requisition_Requested2(0,"+Request.QueryString("insPurchase_Req_id")+",0,0,0,'',0.0,'',0,0,0,0,0,'Q',0)}";
rsInventoryRequested.CursorType = 0;
rsInventoryRequested.CursorLocation = 2;
rsInventoryRequested.LockType = 3;
rsInventoryRequested.Open();

var VID = "";
if (!rsInventoryRequested.EOF) {
	VID = rsInventoryRequested.Fields.Item("insVendor_id").Value;
}
%>
<html>
<head>
	<title>New Inventory Request</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=0,top=0,status=1");
		return ;
	}	   
	
	function Save(){
		if (!CheckDate(document.frm0201.EstimatedDeliveryDate.value)) {
			alert("Invalid Estimated Delivery Date.  Use (mm/dd/yyyy).");
			document.frm0201.EstimatedDeliveryDate.focus();
			return ;
		}
		if (document.frm0201.ClassID.value==0) {
			alert("Select a class.");
			document.frm0201.ListClass.focus();
			return ;
		}
		if (document.frm0201.Quantity.value < 1) {
			alert("Quantity must be greater or equal to 1.");
			document.frm0201.Quantity.focus();
			return ;
		}
		
		<%
		if (VID != "") {
		%>
//		if (document.frm0201.Vendor.value!=document.frm0201.VID.value) {
//			alert("You must select the same vendor for all requests in the same PR.");
//			document.frm0201.Vendor.focus();
//			return ;
//		}
		<%
		}
		%>
		
		document.frm0201.MM_Insert.value="true";
		document.frm0201.submit();
	}
	
	function CalculateTotal(){
		var temp = new Number("0");
		var temp1 = new Number(document.frm0201.Quantity.value);
		var temp2 = new Number(document.frm0201.ListUnitCost.value);
		temp = Math.round(temp1 * temp2 * 100)/100;
		document.frm0201.Total.value= temp.toString();
	}
	
	function SelectVendor(){
		if (document.frm0201.Vendor.length==0) {
			document.frm0201.ListUnitCost.value=0;
			CalculateTotal();
			return ;
		}
		if (document.frm0201.Vendor.length==1) {
			document.frm0201.ListUnitCost.value=document.frm0201.LUC.value;
			CalculateTotal();
			return ;
		} else {
			document.frm0201.ListUnitCost.value=document.frm0201.LUC[document.frm0201.Vendor.selectedIndex].value;
			CalculateTotal();			
			return ;
		}		
	}
	
	function ListUser(){
		if (document.frm0201.UserType.value=="0") {
			openWindow('m014p0201.asp','');
		} else {
			openWindow('m014p0203.asp','');		
		}
	}
	
	function Init(){
		SelectVendor();
		document.frm0201.ListClass.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0201" method="POST" action="<%=MM_editAction%>">
<h5>New Inventory Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Inventory Class:</td>
		<td nowrap>
			<input type="text" name="ClassName" value="<%=Request.QueryString("ClassName")%>" size="60" tabindex="1" accesskey="F" readonly>
			<input type="button" name="ListClass" value="List Class" tabindex="2" onClick="openWindow('m014p01FS.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>','');" class="btnstyle">
		</td>
	</tr>
	<tr>
		<td nowrap>Vendor:</td>
		<td nowrap><select name="Vendor" tabindex="3" onChange="SelectVendor();">
			<%
			while (!rsInventorySupplier.EOF) {
			%>
				<option value="<%=rsInventorySupplier.Fields.Item("insVendor_id").Value%>" <%=((rsInventorySupplier.Fields.Item("bitIsCurrent").Value=="1")?"SELECTED":"")%>><%=rsInventorySupplier.Fields.Item("chvCompany_Name").Value%>
			<%
				rsInventorySupplier.MoveNext;
			}
			rsInventorySupplier.MoveFirst;				
			%>
		</select></td>
	</tr>
	<tr>
		<td nowrap valign="top">Description:</td>
		<td nowrap valign="top"><textarea name="Description" rows="10" cols="65" tabindex="4"></textarea></td>
	</tr>
	<tr>
		<td nowrap>List Unit Cost:</td>
		<td nowrap>$<input type="text" name="ListUnitCost" size="8" tabindex="5" readonly></td>
	</tr>
	<tr>
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="Quantity" size="6" tabindex="6" value="1" onKeypress="AllowNumericOnly();" onChange="CalculateTotal();"></td>
	</tr>
	<tr>
		<td nowrap>Total:</td>
		<td nowrap>$<input type="text" name="Total" size="6" tabindex="7" value="0" readonly></td>
	</tr>
	<tr>
		<td nowrap>For User:</td>
		<td nowrap>
			<select name="UserType" tabindex="8" onChange="document.frm0201.UserName.value='';document.frm0201.UserID.value=0;">
				<option value="0">Client
				<option value="1">Institution
			</select>
			<input type="text" name="UserName" size="30" tabindex="9" readonly>
			<input type="button" value="List" onClick="ListUser();" tabindex="10" class="btnstyle">
		</td>
	</tr>
    <tr>
		<td nowrap>ETA:</td>
		<td nowrap>
			<input type="text" name="EstimatedDeliveryDate" maxlength="10" size="11" tabindex="11" accesskey="L" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="12" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="13" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="UserID" value="0">
<input type="hidden" name="MM_Insert" value="false">
<input type="hidden" name="ClassID" value="<%=Request.QueryString("ClassID")%>">
<%
while (!rsInventorySupplier.EOF) {
%>
<input type="hidden" name="LUC" value="<%=rsInventorySupplier.Fields.Item("fltList_Unit_Cost").Value%>">
<%
	rsInventorySupplier.MoveNext;
}
%>

<input type="hidden" name="VID" value="<%=VID%>">
</form>
</body>
</html>
<%
rsInventorySupplier.Close();
%>
