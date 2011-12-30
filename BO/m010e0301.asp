<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

switch (String(Request("MM_action"))) {
	case "GetInventory":
		var rsInventory = Server.CreateObject("ADODB.Recordset");
		rsInventory.ActiveConnection = MM_cnnASP02_STRING;
		rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory_barcode(1,0,'',1," + Request.Form("InventoryID") + ",0)}";
		rsInventory.CursorType = 0;
		rsInventory.CursorLocation = 2;
		rsInventory.LockType = 3;
		rsInventory.Open();		
	break;
	case "Update":
		var Comments = String(Request.Form("Comments")).replace(/'/g, "''");		
		var SoldPrice = ((Request.Form("SoldPrice")=="")?"0":Request.Form("SoldPrice"));
		var rsInventorySold = Server.CreateObject("ADODB.Recordset");
		rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
		rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold("+Request.QueryString("intBO_Eqp_Sold_id")+","+Request.QueryString("intBuyout_req_id")+","+Request.Form("hInventoryID")+","+SoldPrice+",'"+Request.Form("DateReturned")+"',"+Request.Form("ReturnedBy")+","+Request.Form("ReturnCondition")+",'"+Comments+"',0,'E',0)}";		
		rsInventorySold.CursorType = 0;
		rsInventorySold.CursorLocation = 2;
		rsInventorySold.LockType = 3;
		rsInventorySold.Open();
		
		//Trigger to return the inventory status to "In Stock"		
		if (String(Request.Form("ReturnEquipment"))=="true") {
			var SetInventoryStatus = Server.CreateObject("ADODB.Recordset");
			SetInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
			SetInventoryStatus.Source = "{call dbo.cp_update_eqpivtry_status("+Request.Form("InventoryID")+",1,0)}";		
			SetInventoryStatus.CursorType = 0;
			SetInventoryStatus.CursorLocation = 2;
			SetInventoryStatus.LockType = 3;
			SetInventoryStatus.Open();

			//Trigger to change Buyout Status if all inventory has been returned
			var trigger = true;
			var rsInventorySold = Server.CreateObject("ADODB.Recordset");
			rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
			rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold(0,"+Request.QueryString("intBuyout_req_id")+",0,0.0,'',0,0,'',0,'Q',0)}";
			rsInventorySold.CursorType = 0;
			rsInventorySold.CursorLocation = 2;
			rsInventorySold.LockType = 3;
			rsInventorySold.Open();
			if (rsInventorySold.EOF) {
				trigger = false;
			} else {
				while ((!rsInventorySold.EOF) && trigger) {
					if (rsInventorySold.Fields.Item("insCurrent_Status").Value!=1) trigger = false;
					rsInventorySold.MoveNext();
				}
			}
			
			if (trigger) {
				var SetBuyoutStatus = Server.CreateObject("ADODB.Recordset");
				SetBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
				SetBuyoutStatus.Source = "{call dbo.cp_update_buyout_status("+Request.QueryString("intBuyout_req_id")+",5,0)}";
				SetBuyoutStatus.CursorType = 0;
				SetBuyoutStatus.CursorLocation = 2;
				SetBuyoutStatus.LockType = 3;
				SetBuyoutStatus.Open();
			}			
		}				
		Response.Redirect("UpdateSuccessful.asp?page=m010q0301.asp&intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));
	break;
	case "undefined":
	break;
}

var rsInventorySold = Server.CreateObject("ADODB.Recordset");
rsInventorySold.ActiveConnection = MM_cnnASP02_STRING;
rsInventorySold.Source = "{call dbo.cp_buyout_eqp_sold("+Request.QueryString("intBO_Eqp_Sold_id")+","+Request.QueryString("intBuyout_req_id")+",0,0,'',0,0,'',1,'Q',0)}";
rsInventorySold.CursorType = 0;
rsInventorySold.CursorLocation = 2;
rsInventorySold.LockType = 3;
rsInventorySold.Open();

var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
//if (String(Request("MM_action"))=="GetInventory") {
//	rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory(1,0,'',1," + Request.Form("InventoryID") + ",0)}";
//} else {
rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory(1,0,'',1," + rsInventorySold.Fields.Item("intEquip_set_id").Value + ",0)}";
//}
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();		

var rsStatus = Server.CreateObject("ADODB.Recordset");
rsStatus.ActiveConnection = MM_cnnASP02_STRING;
rsStatus.Source = "{call dbo.cp_ASP_lkup(36)}";
rsStatus.CursorType = 0;
rsStatus.CursorLocation = 2;
rsStatus.LockType = 3;
rsStatus.Open();	

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>Update Equipment Sold</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save1(0);
			break;
		   	case 76 :
				//alert("L");
				window.location.href='m010q0301.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>';
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
		if (document.frm0301.InventoryID.value > 0) temp=window.showModalDialog("m010pop.asp?InventoryID="+document.frm0301.InventoryID.value,"","dialogHeight: 200px; dialogWidth: 375px; dialogTop: px; dialogLeft: px; edge: Sunken; center: Yes; help: No; resizable: No; status: No;");	
	}
		
	function Save1(return_eq){
		if (document.frm0301.InventoryID.value==0) {
			alert("Select a Inventory.");
			document.frm0301.InventoryID.focus();
			return ;
		}
		if (return_eq==1) {
			document.frm0301.ReturnedBy.value="<%=Session("insStaff_id")%>";
			document.frm0301.DateReturned.value="<%=CurrentDate()%>";
			document.frm0301.ReturnEquipment.value="true";		
		}
		
		if (Trim(document.frm0301.SoldPrice.value)=="") {
			document.frm0301.SoldPrice.value = "<%=rsInventory.Fields.Item("fltList_Unit_Cost").Value%>";
		}
						
		document.frm0301.MM_action.value="Update";
		document.frm0301.submit();
	}
	
	function Init(){
		if (document.frm0301.ReturnedBy.value > 0) document.frm0301.Return.disabled = true;
<%
if (String(Request("MM_action"))=="GetInventory") {
	if (!rsInventory.EOF) {
%>
		document.frm0301.InventoryName.value="<%=FilterQuotes(rsInventory.Fields.Item("chvInventory_Name").Value)%>";
		document.frm0301.hInventoryID.value="<%=rsInventory.Fields.Item("intEquip_Set_id").Value%>";		
		document.frm0301.InventoryStatus.value="<%=(rsInventory.Fields.Item("insCurrent_Status").Value)%>";
		document.frm0301.Vendor.value="<%=FilterQuotes(rsInventory.Fields.Item("chvVendor_Name").Value)%>";
		document.frm0301.ModelNumber.value="<%=(rsInventory.Fields.Item("chvModel_Number").Value)%>";
		document.frm0301.SerialNumber.value="<%=(rsInventory.Fields.Item("chvSerial_Number").Value)%>";
		document.frm0301.PRNumber.value="<%=ZeroPadFormat(rsInventory.Fields.Item("intRequisition_no").Value,8)%>";
		document.frm0301.EquipmentCost.value="<%=FormatCurrency(rsInventory.Fields.Item("fltList_Unit_Cost").Value)%>";
<%
		if (((rsInventory.Fields.Item("insCurrent_Status").Value)!="1") && ((rsInventory.Fields.Item("insCurrent_Status").Value)!="6")) {
%>
		alert("This equipment is not available for sale.");
		document.frm0301.Save.disabled=true;
		document.frm0301.Return.disabled=true;
<%
		}
	} else {
%>
		alert("Equipment not found.");
		document.frm0301.Save.disabled=true;
		document.frm0301.Return.disabled=true;
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
<h5>Equipment Sold</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Inventory ID:</td>
		<td nowrap colspan="3">
			<input type="text" name="InventoryID" readonly size="12" value="<%=rsInventory.Fields.Item("intBar_Code_no").Value%>" tabindex="1" accesskey="F" onKeypress="AllowNumericOnly();">
			<input type="button" value="Check Inventory" tabindex="2" onClick="CheckInventory();" class="btnstyle">
			<input type="button" name="ViewAccessory" value="View Accessory" tabindex="3" onClick="ViewAcc();" class="btnstyle">
		</td>
	</tr>
	<tr>
		<td nowrap>Inventory Name:</td>
		<td nowrap colspan="3"><input type="text" name="InventoryName" size="60" value="<%=rsInventory.Fields.Item("chvInventory_Name").Value%>" tabindex="4" readonly></td>
	</tr>	
	<tr>
		<td nowrap>Inventory Status:</td>
		<td nowrap colspan="3"><select name="InventoryStatus" tabindex="5">
		<% 
		while (!rsStatus.EOF) { 			
		%>
			<option value="<%=(rsStatus.Fields.Item("insEquip_status_id").Value)%>" <%=((rsStatus.Fields.Item("insEquip_status_id").Value==rsInventory.Fields.Item("insCurrent_Status").Value)?"SELECTED":"")%>><%=(rsStatus.Fields.Item("chvStatusDesc").Value)%>
		<% 
			rsStatus.MoveNext();
		} 
		%>		
		</select></td>
	</tr>
	<tr>
		<td nowrap>Vendor:</td>
		<td nowrap colspan="3"><input type="textbox" name="Vendor" size="40" value="<%=rsInventory.Fields.Item("chvVendor_Name").Value%>" tabindex="6" readonly></td>
	</tr>
	<tr>
		<td nowrap>Model Number:</td>
		<td nowrap><input type="text" name="ModelNumber" size="12" tabindex="7" value="<%=rsInventory.Fields.Item("chvModel_Number").Value%>" readonly></td>
		<td nowrap>Serial Number:</td>
		<td nowrap><input type="text" name="SerialNumber" size="12" tabindex="8" value="<%=rsInventory.Fields.Item("chvSerial_Number").Value%>" readonly></td>
	</tr>
	<tr>
		<td nowrap>PR Number:</td>
		<td nowrap><input type="text" name="PRNumber" size="12" tabindex="9" value="<%=ZeroPadFormat(rsInventory.Fields.Item("intRequisition_no").Value,8)%>" readonly></td>
		<td nowrap>Equipment Cost:</td>
		<td nowrap><input type="text" name="EquipmentCost" size="12" tabindex="10" value="<%=FormatCurrency(rsInventory.Fields.Item("fltList_Unit_Cost").Value)%>" readonly></td>
	</tr>
	<tr>
		<td nowrap>Date Returned:</td>
		<td nowrap>
			<input type="text" name="DateReturned" size="11" tabindex="11" value="<%=FilterDate(rsInventorySold.Fields.Item("dtsDate_Returned").Value)%>" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>		
		<td nowrap>Sold Price:</td>
		<td nowrap>$<input type="text" name="SoldPrice" size="11" value="<%=(((Request.Form("SoldPrice")=="")||(String(Request.Form("SoldPrice"))=="undefined"))?rsInventorySold.Fields.Item("fltEqp_Sold_Price").Value:Request.Form("SoldPrice"))%>" tabindex="12" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td nowrap>Returned By:</td>
		<td nowrap><select name="ReturnedBy" tabindex="13">
			<option value="0">(none)
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==rsInventorySold.Fields.Item("insRtned_by_id"))?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		%>		
		</select></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top" colspan="3"><textarea name="Comments" rows="5" cols="65" tabindex="14" accesskey="L"><%=rsInventorySold.Fields.Item("chvComments").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" name="Save" value="Save" onClick="Save1(0);" tabindex="15" class="btnstyle"></td>
		<td><input type="button" name="Return" value="Return This Equipment" onClick="Save1(1);" tabindex="16" class="btnstyle"></td>		
		<td><input type="button" value="Close" onClick="window.location.href='m010q0301.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>';" tabindex="17" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="">
<input type="hidden" name="ClassID" value="<%if ((!rsInventory.EOF) && (rsInventory.Fields.Item("insEquip_Class_id").Value > 0)) Response.Write(rsInventory.Fields.Item("insEquip_Class_id").Value)%>">
<input type="hidden" name="ReturnEquipment" value="false">
<input type="hidden" name="ReturnCondition" value="1">
<input type="hidden" name="hInventoryID" value="<%=rsInventory.Fields.Item("intEquip_Set_id").Value%>">
</form>
</body>
</html>
<%
rsInventorySold.Close();
rsInventory.Close();
rsStaff.Close();
rsStatus.Close();
%>