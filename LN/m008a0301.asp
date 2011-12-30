<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

//InventoryID on the screen is the barcode ID that the user enters.
//The real intEquip_set_id is hInventoryID
var ClassID = 0;
switch (String(Request("MM_action"))) {
	case "GetInventory":
		var rsInventory = Server.CreateObject("ADODB.Recordset");
		rsInventory.ActiveConnection = MM_cnnASP02_STRING;
		rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory_barcode(1,0,'',1," + Request.Form("InventoryID") + ",0)}";
		rsInventory.CursorType = 0;
		rsInventory.CursorLocation = 2;
		rsInventory.LockType = 3;
		rsInventory.Open();
		if (!rsInventory.EOF) ClassID = rsInventory.Fields.Item("insEquip_Class_id").Value;
	break;
	case "Insert":
		var Comments = String(Request.Form("Comments")).replace(/'/g, "''");		
		var rsInventoryLoan = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoan.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoan.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_req_id")+","+Request.Form("hInventoryID")+",'',0,0,'"+Request.Form("DateProcessed")+"','"+Comments+"',0,'A',0)}";
		rsInventoryLoan.CursorType = 0;
		rsInventoryLoan.CursorLocation = 2;
		rsInventoryLoan.LockType = 3;
		rsInventoryLoan.Open();
		Response.Redirect("AddDeleteSuccessful.asp?action=Add");	
	break;
}

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_ASP_lkup(36)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();	
%>
<html>
<head>
	<title>New Equipment Loaned</title>
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
	
	function CheckInventory(){
		if (document.frm0301.InventoryID.value==0) {
			alert("Select a Inventory.");
			document.frm0301.InventoryID.focus();
			return ;
		}
		document.frm0301.MM_action.value="GetInventory";
		document.frm0301.submit();
	}

	function ViewAcc(){	
		if (document.frm0301.ClassID.value > 0) temp=window.showModalDialog("m008pop.asp?ClassID="+document.frm0301.ClassID.value,"","dialogHeight: 200px; dialogWidth: 375px; dialogTop: px; dialogLeft: px; edge: Sunken; center: Yes; help: No; resizable: No; status: No;");	
	}
	
	function Save(){	
		if (document.frm0301.InventoryID.value==0) {
			alert("Select a Inventory.");
			document.frm0301.InventoryID.focus();
			return ;
		}
		if ((!CheckDate(document.frm0301.DateProcessed.value)) || (Trim(document.frm0301.DateProcessed.value)=="")){
			alert("Invalid Date Processed.");
			document.frm0301.DateProcessed.focus();
			return ;
		}		
		document.frm0301.MM_action.value="Insert";
		document.frm0301.submit();
	}
	
	function Init(){
<%
if (String(Request("MM_action"))=="GetInventory") {
	if (!rsInventory.EOF) {
%>
		document.frm0301.InventoryName.value="<%=FilterQuotes(rsInventory.Fields.Item("chvInventory_Name").Value)%>";
		document.frm0301.hInventoryID.value="<%=(rsInventory.Fields.Item("intEquip_Set_id").Value)%>";
		document.frm0301.InventoryStatus.value="<%=(rsInventory.Fields.Item("insCurrent_Status").Value)%>";
		document.frm0301.Vendor.value="<%=FilterQuotes(rsInventory.Fields.Item("chvVendor_Name").Value)%>";
		document.frm0301.ModelNumber.value="<%=(rsInventory.Fields.Item("chvModel_Number").Value)%>";
		document.frm0301.SerialNumber.value="<%=(rsInventory.Fields.Item("chvSerial_Number").Value)%>";
		document.frm0301.PurchaseRequisitionNumber.value="<%=(rsInventory.Fields.Item("intRequisition_no").Value)%>";
		document.frm0301.EquipmentCost.value="<%=FormatCurrency(rsInventory.Fields.Item("fltList_Unit_Cost").Value)%>";
<%
		if (rsInventory.Fields.Item("insCurrent_Status").Value!="1") {
%>
		alert("This equipment is not available for loan.");
		document.frm0301.btnSave.disabled = true;
<%
		}
	} else {
%>
		alert("Equipment not found.");
		document.frm0301.btnSave.disabled = true;
<%
	}
}
%>
		document.frm0301.InventoryID.focus();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0301" method="POST" action="<%=MM_editAction%>">
<h5>New Equipment Loaned</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Inventory ID:</td>
		<td nowrap>
			<input type="text" name="InventoryID" size="12" value="<%=Request("InventoryID")%>" tabindex="1" accesskey="F" onKeypress="AllowNumericOnly();">
			<input type="button" value="Check Inventory" tabindex="2" onClick="CheckInventory();" class="btnstyle">
			<input type="button" name="ViewAccessory" value="View Accessory" tabindex="3" onClick="ViewAcc();" class="btnstyle">			
		</td>
	</tr>
	<tr>
		<td nowrap>Inventory Name:</td>
		<td nowrap><input type="text" name="InventoryName" size="60" tabindex="4" readonly></td>
	</tr>	
	<tr>
		<td nowrap>Inventory Status:</td>
		<td nowrap><select name="InventoryStatus" tabindex="5">
			<% 
			while (!rsStatus.EOF) { 			
			%>
				<option value="<%=(rsStatus.Fields.Item("insEquip_status_id").Value)%>"><%=(rsStatus.Fields.Item("chvStatusDesc").Value)%>
			<% 
				rsStatus.MoveNext();
			} 
			%>		
		</select></td>
	</tr>
	<tr>
		<td nowrap>Vendor:</td>
		<td nowrap><input type="textbox" name="Vendor" size="40" tabindex="6" readonly></td>
	</tr>
	<tr>
		<td nowrap>Model Number:</td>
		<td nowrap><input type="text" name="ModelNumber" size="12" tabindex="7" readonly></td>
	</tr>
	<tr>
		<td nowrap>Serial Number:</td>
		<td nowrap><input type="text" name="SerialNumber" size="12" tabindex="8" readonly></td>
	</tr>
	<tr>
		<td nowrap>PR Number:</td>
		<td nowrap><input type="text" name="PurchaseRequisitionNumber" size="12" tabindex="9" readonly></td>
	</tr>
	<tr>
		<td nowrap>Equipment Cost:</td>
		<td nowrap><input type="text" name="EquipmentCost" size="12" tabindex="10" readonly></td>
	</tr>
	<tr>
		<td nowrap>Date Processed:</td>
		<td nowrap>
			<input type="text" name="DateProcessed" size="11" tabindex="11" value="<%=((String(Request.Form("DateProcessed"))=="undefined")?CurrentDate():Request.Form("DateProcessed"))%>" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>		
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" rows="5" cols="65" tabindex="12" accesskey="L"><%=Request.Form("Comments")%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" name="btnSave" value="Save" onClick="Save();" tabindex="13" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="14" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action">
<input type="hidden" name="ClassID" value="<%=ClassID%>">
<input type="hidden" name="hInventoryID">
</form>
</body>
</html>
<%
rsStatus.Close();
%>