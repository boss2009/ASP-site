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
		var rsInventoryLoan = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoan.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoan.Source = "{call dbo.cp_eqp_loaned("+Request.QueryString("intEqp_Loaned_Id")+","+Request.QueryString("intLoan_req_id")+","+Request.Form("hInventoryID")+",'"+Request.Form("DateReturned")+"',"+Request.Form("ReturnedBy")+","+Request.Form("ReturnCondition")+",'"+Request.Form("DateProcessed")+"','"+Comments+"',0,'E',0)}";		
		rsInventoryLoan.CursorType = 0;
		rsInventoryLoan.CursorLocation = 2;
		rsInventoryLoan.LockType = 3;
		rsInventoryLoan.Open();

		//Trigger to return the inventory status to "In Stock"		
		if (String(Request.Form("ReturnEquipment"))=="true") {
			var SetInventoryStatus = Server.CreateObject("ADODB.Recordset");
			SetInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
			SetInventoryStatus.Source = "{call dbo.cp_update_eqpivtry_status("+Request.Form("InventoryID")+",1,0)}";		
			SetInventoryStatus.CursorType = 0;
			SetInventoryStatus.CursorLocation = 2;
			SetInventoryStatus.LockType = 3;
			SetInventoryStatus.Open();

			//Trigger to change Loan Status if all inventory has been returned
			var trigger = true;
			var rsInventoryLoan = Server.CreateObject("ADODB.Recordset");
			rsInventoryLoan.ActiveConnection = MM_cnnASP02_STRING;
			rsInventoryLoan.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_Req_id")+",0,'',0,0,'','',0,'Q',0)}";
			rsInventoryLoan.CursorType = 0;
			rsInventoryLoan.CursorLocation = 2;
			rsInventoryLoan.LockType = 3;
			rsInventoryLoan.Open();
			if (rsInventoryLoan.EOF) {
				trigger = false;
			} else {
				while ((!rsInventoryLoan.EOF) && trigger) {
					if (rsInventoryLoan.Fields.Item("insCurrent_Status").Value!=1) trigger = false;
					rsInventoryLoan.MoveNext();
				}
			}
			
			if (trigger) {
				var SetLoanStatus = Server.CreateObject("ADODB.Recordset");
				SetLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
				SetLoanStatus.Source = "update tbl_loan_request set insLoan_Status_id = 4 where intLoan_req_id = " + Request.QueryString("intLoan_req_id");
				SetLoanStatus.CursorType = 0;
				SetLoanStatus.CursorLocation = 2;
				SetLoanStatus.LockType = 3;
				SetLoanStatus.Open();
			}			
		}		
		Response.Redirect("UpdateSuccessful.asp?page=m008q0301.asp&intLoan_req_id="+Request.QueryString("intLoan_req_id"));
	break;
	case "undefined":
		var rsInventoryLoan = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoan.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoan.Source = "{call dbo.cp_eqp_loaned("+Request.QueryString("intEqp_Loaned_Id")+",0,0,'',0,0,'','',1,'Q',0)}";
		rsInventoryLoan.CursorType = 0;
		rsInventoryLoan.CursorLocation = 2;
		rsInventoryLoan.LockType = 3;
		rsInventoryLoan.Open();
	
		var rsInventory = Server.CreateObject("ADODB.Recordset");
		rsInventory.ActiveConnection = MM_cnnASP02_STRING;
		rsInventory.Source = "{call dbo.cp_Get_EqCls_Inventory_barcode(1,0,'',1," + rsInventoryLoan.Fields.Item("intBar_Code_no").Value + ",0)}";
		rsInventory.CursorType = 0;
		rsInventory.CursorLocation = 2;
		rsInventory.LockType = 3;
		rsInventory.Open();			
	break;
}

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
	<title>Update Equipment Loaned</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save(0);
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
		
	function Save(return_eq){
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
		document.frm0301.MM_action.value="Update";
		document.frm0301.submit();
	}
	
	function Init(){
		if (document.frm0301.ReturnedBy.value > 0) document.frm0301.Return.disabled = true;
<%
var InventoryID = "";
var hInventoryID = "";
var InventoryName = "";
var InventoryStatus = 4;
var Vendor = "";
var ModelNumber = "";
var SerialNumber = "";
var PurchaseRequisitionNumber = "";
var EquipmentCost = 0;
var DateProcessed = "";
var DateReturned = "";
var ReturnedBy = 0;
var ReturnComplete = 0;
var Comments = "";

switch (String(Request.Form("MM_action"))) {
	case "GetInventory":
		InventoryID = Request.Form("InventoryID");	
		if (!rsInventory.EOF) {
			InventoryName = FilterQuotes(rsInventory.Fields.Item("chvInventory_Name").Value);
			hInventoryID = rsInventory.Fields.Item("intEquip_Set_id").Value;
			InventoryStatus = rsInventory.Fields.Item("insCurrent_Status").Value;
			Vendor = FilterQuotes(rsInventory.Fields.Item("chvVendor_Name").Value);
			ModelNumber = rsInventory.Fields.Item("chvModel_Number").Value;
			SerialNumber = rsInventory.Fields.Item("chvSerial_Number").Value;
			PurchaseRequisitionNumber = ZeroPadFormat(rsInventory.Fields.Item("intRequisition_no").Value,8);
			EquipmentCost = FormatCurrency(rsInventory.Fields.Item("fltList_Unit_Cost").Value);		
			if (((rsInventory.Fields.Item("insCurrent_Status").Value)!="1") && ((rsInventory.Fields.Item("insCurrent_Status").Value)!="6")) {
%>
			alert("This equipment is not available for loan.");
			document.frm0301.btnSave.disabled=true;
<%
			}
		} else {
%>
			alert("Equipment not found.");
			document.frm0301.btnSave.disabled=true;
<%
		}
	break;
	case "undefined":
		if (!rsInventory.EOF) {
			InventoryID = rsInventoryLoan.Fields.Item("intBar_Code_no").Value;
			InventoryName = FilterQuotes(rsInventoryLoan.Fields.Item("chvInventory_Name").Value);
			hInventoryID = rsInventory.Fields.Item("intEquip_Set_id").Value;
			InventoryStatus = rsInventoryLoan.Fields.Item("insCurrent_Status").Value;
			Vendor = FilterQuotes(rsInventoryLoan.Fields.Item("chvVendor_Name").Value);
			ModelNumber = rsInventoryLoan.Fields.Item("chvModel_Number").Value;
			SerialNumber = rsInventoryLoan.Fields.Item("chvSerial_Number").Value;
			PurchaseRequisitionNumber = ZeroPadFormat(rsInventoryLoan.Fields.Item("intRequisition_no").Value,8);
			EquipmentCost = FormatCurrency(rsInventoryLoan.Fields.Item("fltList_Unit_Cost").Value);
			DateProcessed = FilterDate(rsInventoryLoan.Fields.Item("dtsDate_Shipped").Value);
			DateReturned = FilterDate(rsInventoryLoan.Fields.Item("dtsDate_Returned").Value);
			ReturnedBy = rsInventoryLoan.Fields.Item("insReturned_by_id").Value;
			ReturnComplete = rsInventoryLoan.Fields.Item("bitRtn_Complete").Value;
			Comments = 	FilterDate(rsInventoryLoan.Fields.Item("chvComments").Value);		
		}	
	break;
}
%>
			document.frm0301.InventoryID.focus();		
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0301" method="POST" action="<%=MM_editAction%>">
<h5>Equipment Loaned</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Inventory ID:</td>
		<td nowrap colspan="3">
			<input type="text" name="InventoryID" readonly value="<%=InventoryID%>" size="12" tabindex="1" accesskey="F" onKeypress="AllowNumericOnly();">			
			<input type="button" value="Check Inventory" tabindex="2" onClick="CheckInventory();" class="btnstyle">
			<input type="button" name="ViewAccessory" value="View Accessory" tabindex="3" onClick="ViewAcc();" class="btnstyle">			
		</td>
	</tr>
	<tr>
		<td nowrap>Inventory Name:</td>
		<td nowrap colspan="3"><input type="text" name="InventoryName" value="<%=InventoryName%>" size="50" tabindex="4" readonly></td>
	</tr>	
	<tr>
		<td nowrap>Inventory Status:</td>
		<td nowrap colspan="3">
			<select name="InventoryStatus" tabindex="5" disabled>
			<% 
			while (!rsStatus.EOF) { 			
			%>
				<option value="<%=(rsStatus.Fields.Item("insEquip_status_id").Value)%>" <%=((rsStatus.Fields.Item("insEquip_status_id").Value==InventoryStatus)?"SELECTED":"")%>><%=(rsStatus.Fields.Item("chvStatusDesc").Value)%>
			<% 
				rsStatus.MoveNext();
			} 
			%>		
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>Vendor:</td>
		<td nowrap colspan="3"><input type="textbox" name="Vendor" value="<%=Vendor%>" size="40" tabindex="6" readonly></td>
	</tr>
	<tr>
		<td nowrap>Model Number:</td>
		<td nowrap><input type="text" name="ModelNumber" value="<%=ModelNumber%>" size="12" tabindex="7" readonly></td>
		<td nowrap>Serial Number:</td>
		<td nowrap><input type="text" name="SerialNumber" value="<%=SerialNumber%>" size="12" tabindex="8" readonly></td>
	</tr>
	<tr>
		<td nowrap>PR Number:</td>
		<td nowrap><input type="text" name="PurchaseRequisitionNumber" value="<%=ZeroPadFormat(PurchaseRequisitionNumber)%>" size="12" tabindex="9" readonly></td>
		<td nowrap>Equipment Cost:</td>
		<td nowrap><input type="text" name="EquipmentCost" value="<%=FormatCurrency(EquipmentCost)%>" size="12" tabindex="10" readonly></td>
	</tr>
	<tr>
		<td nowrap>Date Processed:</td>
		<td nowrap>
			<input type="text" name="DateProcessed" value="<%=FilterDate(DateProcessed)%>" size="11" tabindex="11" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>		
		<td nowrap>Date Returned:</td>
		<td nowrap>
			<input type="text" name="DateReturned" value="<%=FilterDate(DateReturned)%>" size="11" tabindex="12" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>		
	</tr>
	<tr>
		<td nowrap>Returned By:</td>
		<td nowrap><select name="ReturnedBy" tabindex="13">
			<option value="0">(none)
		<%
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==ReturnedBy)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%> 
		<%
			rsStaff.MoveNext();
		}
		%>		
		</select></td>
		<td nowrap>Return Condition:</td>
		<td nowrap><select name="ReturnCondition" tabindex="14">
			<option value="0" <%=((ReturnComplete=="0")?"SELECTED":"")%>>Incomplete
			<option value="1" <%=((ReturnComplete=="1")?"SELECTED":"")%>>Complete
		</select></td>
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap colspan="3"><textarea name="Comments" rows="5" cols="65" tabindex="15" accesskey="L"><%=Comments%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" name="btnSave" value="Save" onClick="Save(0);" tabindex="16" class="btnstyle"></td>
		<td><input type="button" name="Return" value="Return This Equipment" onClick="Save(1);" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m008q0301.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>'" tabindex="17" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="">
<input type="hidden" name="ReturnEquipment" value="false">
<input type="hidden" name="ClassID" value="<%=(((!rsInventory.EOF) && (rsInventory.Fields.Item("insEquip_Class_id").Value > 0))?rsInventory.Fields.Item("insEquip_Class_id").Value:"")%>">
<input type="hidden" name="hInventoryID" value="<%=hInventoryID%>">
</form>
</body>
</html>
<%
rsInventory.Close();
rsStaff.Close();
rsStatus.Close();
%>