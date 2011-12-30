<!--------------------------------------------------------------------------
* File Name: m014e0201.asp
* Title: Edit Inventory Request
* Main SP: cp_purchase_requisition_requested
* Description: This page updates inventory requested.
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

if (String(Request("MM_update")) == "true") {	
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var ClassID = ((String(Request.Form("ClassID"))!="undefined")?Request.Form("ClassID"):0);
	var Quantity = ((String(Request.Form("Quantity"))!="")?Request.Form("Quantity"):0);
	var ListUnitCost = ((String(Request.Form("ListUnitCost"))!="")?ListUnitCost=Request.Form("ListUnitCost"):0);
	var EstimatedDeliveryDate = ((String(Request.Form("EstimatedDeliveryDate"))!="undefined")?Request.Form("EstimatedDeliveryDate"):"");	
	var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
	rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
	rsInventoryRequest.Source = "{call dbo.cp_Purchase_Requisition_Requested2("+Request.QueryString("insRqst_requested_id")+","+Request.QueryString("insPurchase_Req_id")+","+ClassID+",1,"+Quantity+",'"+Description+"',"+ListUnitCost+",'"+EstimatedDeliveryDate+"',"+Request.Form("UserID")+","+Request.Form("UserID")+","+Request.Form("UserType")+","+Request.Form("Vendor")+",0,'E',0)}";
	rsInventoryRequest.CursorType = 0;
	rsInventoryRequest.CursorLocation = 2;
	rsInventoryRequest.LockType = 3;
	rsInventoryRequest.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m014q0201.asp&insPurchase_Req_id="+Request.QueryString("insPurchase_Req_id"));
}

var rsInventoryRequest = Server.CreateObject("ADODB.Recordset");
rsInventoryRequest.ActiveConnection = MM_cnnASP02_STRING;
rsInventoryRequest.Source = "{call dbo.cp_Purchase_Requisition_Requested2("+Request.QueryString("insRqst_requested_id")+",0,0,0,0,'',0.0,'',0,0,0,0,1,'Q',0)}";
rsInventoryRequest.CursorType = 0;
rsInventoryRequest.CursorLocation = 2;
rsInventoryRequest.LockType = 3;
rsInventoryRequest.Open();

var ClassID = ((String(Request.QueryString("ClassID"))=="undefined")?rsInventoryRequest.Fields.Item("insClass_bundle_id").Value:Request.QueryString("ClassID"));

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
	<title>Update Inventory Request</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language=Javascript src="../js/MyFunctions.js"></script>
	<script FOR=document event="onkeyup()" language="JavaScript">
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
		document.frm0201.MM_update.value="true";
		document.frm0201.submit();
	}

	function ListUser(){
		if (document.frm0201.UserType.value=="0") {
			openWindow('m014p0201.asp','');
		} else {
			openWindow('m014p0203.asp','');		
		}
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

	function Init(){
		SelectVendor();
		document.frm0201.ClassName.focus();
	}	
	</script>
</head>
<body onLoad="Init();">
<form name="frm0201" method="POST" action="<%=MM_editAction%>">
<h5>Update Inventory Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Inventory Class:</td>
		<td nowrap> 
			<input type="text" name="ClassName" size="40" tabindex="1" accesskey="F" readonly value="<%=((String(Request.QueryString("ClassName"))=="undefined")?rsInventoryRequest.Fields.Item("chvClass_Bundle_Name").Value:Request.QueryString("ClassName"))%>">
			<input type="button" name="ListClass" value="List Class" tabindex="2" disabled onClick="openWindow('m014p01FS.asp?insPurchase_Req_id=<%=Request.QueryString("insPurchase_Req_id")%>&insRqst_requested_id=<%=Request.QueryString("insRqst_requested_id")%>','');" class="btnstyle">
		</td>
    </tr>
	<tr>
		<td nowrap>Vendor:</td>
		<td nowrap><select name="Vendor" tabindex="3" onChange="SelectVendor();">
				<option value="0">(none)
			<%
			while (!rsInventorySupplier.EOF) {
			%>
				<option value="<%=rsInventorySupplier.Fields.Item("insVendor_id").Value%>" <%=((rsInventorySupplier.Fields.Item("insVendor_id").Value==rsInventoryRequest.Fields.Item("insVendor_id").Value)?"SELECTED":"")%>><%=rsInventorySupplier.Fields.Item("chvCompany_Name").Value%>
			<%
				rsInventorySupplier.MoveNext;
			}
			rsInventorySupplier.MoveFirst;
			%>		
		</select></td>
	</tr>
    <tr> 
		<td valign="top">Description:</td>
		<td valign="top"><textarea name="Description" rows="5" cols="65" tabindex="4"><%=(rsInventoryRequest.Fields.Item("chvDescription").Value)%></textarea></td>
    </tr>
    <tr> 
		<td nowrap>List Unit Cost:</td>
		<td nowrap>$<input type="text" name="ListUnitCost" size="8" tabindex="5" readonly value="<%=((String(Request.QueryString("ListUnitCost"))=="undefined")?rsInventoryRequest.Fields.Item("fltPR_request_List_Unit_Cost").Value:Request.QueryString("ListUnitCost"))%>"></td>
    </tr>
    <tr> 
		<td nowrap>Quantity:</td>
		<td nowrap><input type="text" name="Quantity" size="6" tabindex="6" onKeypress="AllowNumericOnly();" readonly onChange="CalculateTotal();" value="<%=((String(Request.QueryString("ListUnitCost"))=="undefined")?rsInventoryRequest.Fields.Item("insPR_request_Qty_Ordered").Value:"0")%>"></td>
    </tr>
    <tr> 
		<td nowrap>Total:</td>
		<td nowrap>$<input type="text" name="Total" size="6" tabindex="7" readonly value="<%=((String(Request.QueryString("ListUnitCost"))=="undefined")?rsInventoryRequest.Fields.Item("fltTotal_Cost").Value:"0")%>"></td>
    </tr>
	<tr>
		<td nowrap>For User:</td>
		<td nowrap>
			<select name="UserType" tabindex="8" onChange="document.frm0201.UserName.value='';document.frm0201.UserID.value=0;">
				<option value="0" <%=((rsInventoryRequest.Fields.Item("intFor_Adult_id").Value>0)?"SELECTED":"")%>>Client
				<option value="1" <%=((rsInventoryRequest.Fields.Item("insSchool_id").Value>0)?"SELECTED":"")%>>Institution
			</select>
			<input type="text" name="UserName" value="<%=rsInventoryRequest.Fields.Item("chvClient").Value%>" size="30" tabindex="9" readonly>
			<input type="button" value="List" onClick="ListUser();" readonly tabindex="10" class="btnstyle">
		</td>
	</tr>
    <tr> 
		<td nowrap>ETA:</td>
		<td nowrap> 
			<input type="text" name="EstimatedDeliveryDate" maxlength="10" size="11" readonly tabindex="11" accesskey="L" value="<%=FilterDate(rsInventoryRequest.Fields.Item("dtsEst_delivery_date").Value)%>" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>		
		</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><input type="button" value="Save" onClick="Save();" tabindex="12" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="history.back();" tabindex="13" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="ClassID" value="<%=((String(Request.QueryString("ClassID"))=="undefined")?rsInventoryRequest.Fields.Item("insClass_bundle_id").Value:Request.QueryString("ClassID"))%>">
<input type="hidden" name="UserID" value="<%=(rsInventoryRequest.Fields.Item("insSchool_id").Value+rsInventoryRequest.Fields.Item("intFor_Adult_id").Value)%>">
<input type="hidden" name="MM_update" value="false">
<input type="hidden" name="LUC" value="0">
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
rsInventoryRequest.Close();
rsInventorySupplier.Close();
%>