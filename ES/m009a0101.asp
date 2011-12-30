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
	case "Insert":
		var DateRequested = ((String(Request.Form("DateRequested"))!="undefined")?Request.Form("DateRequested"):"1/1/1900");
		var DateReceived = ((String(Request.Form("DateReceived"))!="undefined")?Request.Form("DateReceived"):"1/1/1900");
		var IsReceived = ((Request.Form("IsReceived")=="on")?"1":"0");
		var cmdInsertEquipmentService = Server.CreateObject("ADODB.Command");
		cmdInsertEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
		cmdInsertEquipmentService.CommandText = "dbo.cp_Insert_EqpSrv_A";
		cmdInsertEquipmentService.CommandType = 4;
		cmdInsertEquipmentService.CommandTimeout = 0;
		cmdInsertEquipmentService.Prepared = true;
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("RETURN_VALUE", 3, 4));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@intEquip_Set_id", 3, 1,1,Request.Form("hInventoryID")));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@insRepair_Status", 2, 1,1,Request.Form("RepairStatus")));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@dtsRequested_date", 200, 1,30,DateRequested));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@bitIs_Received", 2, 1,1,IsReceived));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@dtsReceived_date", 200, 1,30,DateReceived));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@chvRqstNote_Desc", 200, 1,4000,Request.Form("Description")));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@intUser_id", 3, 1,1,Request.Form("UserID")));
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@insUsr_type_id", 2, 1,1,Request.Form("UserType")));		
		cmdInsertEquipmentService.Parameters.Append(cmdInsertEquipmentService.CreateParameter("@intRtnFlag", 3, 2));
		cmdInsertEquipmentService.Execute();	
		Response.Redirect("m009FS3.asp?intEquip_srv_id="+cmdInsertEquipmentService.Parameters.Item("@intRtnFlag").Value);
	break;
}

var rsRepairStatus = Server.CreateObject("ADODB.Recordset");
rsRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
rsRepairStatus.Source = "{call dbo.cp_repair_status(0,'',0,'Q',0)}";
rsRepairStatus.CursorType = 0;
rsRepairStatus.CursorLocation = 2;
rsRepairStatus.LockType = 3;
rsRepairStatus.Open();
%>
<html>
<head>
	<title>New Equipment Service</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save1();
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
		if (document.frm0101.InventoryID.value==0) {
			alert("Select a Inventory.");
			document.frm0101.InventoryID.focus();
			return ;
		}
		document.frm0101.MM_action.value="GetInventory";
		document.frm0101.submit();
	}

	function Save1(){
		if (document.frm0101.InventoryID.value==0) {
			alert("Select a Inventory.");
			document.frm0101.InventoryID.focus();
			return ;
		}
		if (!CheckTextArea(document.frm0101.Description, 4000)) {
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}		
		if (!CheckDate(document.frm0101.DateRequested.value)) {
			alert("Invalid Date Requested.");
			document.frm0101.DateRequested.focus();
			return ;
		}		
		if (!CheckDate(document.frm0101.DateReceived.value)) {
			alert("Invalid Date Received.");
			document.frm0101.DateReceived.focus();
			return ;
		}				
		document.frm0101.MM_action.value="Insert";
		document.frm0101.submit();
	}

	function ChangeReceived() {
		if (document.frm0101.IsReceived.checked) {
			document.frm0101.DateReceived.disabled = false;			
			document.frm0101.DateReceived.value = "<%=CurrentDate()%>";			
		} else {
			document.frm0101.DateReceived.disabled = true;		
			document.frm0101.DateReceived.value = "";
		}		
	}
		
	function Init(){
		ChangeReceived();	
<%
if (String(Request("MM_action"))=="GetInventory") {
	if (!rsInventory.EOF) {
%>
		document.frm0101.InventoryName.value="<%=FilterQuotes(rsInventory.Fields.Item("chvInventory_Name").Value)%>";
		document.frm0101.hInventoryID.value="<%=rsInventory.Fields.Item("intEquip_set_id").Value%>";
		document.frm0101.Vendor.value="<%=FilterQuotes(rsInventory.Fields.Item("chvVendor_Name").Value)%>";
		document.frm0101.ModelNumber.value="<%=(rsInventory.Fields.Item("chvModel_Number").Value)%>";
		document.frm0101.SerialNumber.value="<%=(rsInventory.Fields.Item("chvSerial_Number").Value)%>";
		document.frm0101.PRNumber.value="<%=(rsInventory.Fields.Item("intRequisition_no").Value)%>";
		document.frm0101.UserType.value="<%=rsInventory.Fields.Item("insUser_Type_id").Value%>";
<%		
		if (rsInventory.Fields.Item("insUser_Type_id").Value==4) {
%>
		document.frm0101.UserID.value="<%=rsInventory.Fields.Item("insInstit_User_id").Value%>";		
		document.frm0101.CurrentUser.value="<%=FilterQuotes(rsInventory.Fields.Item("chvInstitUsr_Nm").Value)%>";			
<%
		} else {
%>		
		document.frm0101.UserID.value="<%=rsInventory.Fields.Item("insUser_id").Value%>";		
		document.frm0101.CurrentUser.value="<%=FilterQuotes(rsInventory.Fields.Item("chvIdvUsr_Nm").Value)%>";			
<%
		}
	} else {
%>
		alert("Equipment not found.");
		document.frm0101.Save.disabled=true;
<%
	}
}
%>
		document.frm0101.InventoryID.focus();
	}
	</script>
</head>
<body onLoad="Init();">
<form name="frm0101" method="POST" action="<%=MM_editAction%>">
  <h5>New Equipment Service</h5>
  <hr>
  <table cellpadding="1" cellspacing="1">
    <tr> 
      <td nowrap>Inventory ID:</td>
      <td nowrap> 
        <input type="text" name="InventoryID" tabindex="1" size="12" value="<%=Request("InventoryID")%>" accesskey="F" onKeypress="AllowNumericOnly();">
        <input type="button" value="Check Inventory" tabindex="2" onClick="CheckInventory();" class="btnstyle">
      </td>
    </tr>
    <tr> 
      <td nowrap>Inventory Name:</td>
      <td nowrap><input type="text" name="InventoryName" size="60" tabindex="3" readonly></td>
    </tr>
    <tr> 
      <td nowrap>Vendor:</td>
      <td nowrap><input type="text" name="Vendor" size="40" tabindex="4" readonly></td>
    </tr>
    <tr> 
      <td nowrap>Model Number:</td>
      <td nowrap><input type="text" name="ModelNumber" size="20" tabindex="5" readonly></td>
    </tr>
    <tr> 
      <td nowrap>Serial Number:</td>
      <td nowrap><input type="text" name="SerialNumber" size="20" tabindex="6" readonly></td>
    </tr>
    <tr> 
      <td nowrap>PR Number:</td>
      <td nowrap><input type="text" name="PRNumber" size="12" tabindex="7" readonly></td>
    </tr>
    <tr> 
      <td nowrap>Current User:</td>
      <td nowrap><input type="text" name="CurrentUser" size="30" tabindex="8" readonly></td>
    </tr>
    <tr> 
      <td nowrap>Repair Status:</td>
      <td nowrap>
        <select name="RepairStatus" tabindex="9">
          <% 
		while (!rsRepairStatus.EOF) { 			
		%>
          <option value="<%=(rsRepairStatus.Fields.Item("insEq_Repair_Sts_id").Value)%>" <%=((rsRepairStatus.Fields.Item("insEq_Repair_Sts_id").Value==1)?"SELECTED":"")%>><%=(rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value)%> 
          <% 
			rsRepairStatus.MoveNext();
		} 
		%>
        </select>
      </td>
    </tr>
    <tr> 
      <td nowrap>Date Requested:</td>
      <td nowrap> 
        <input type="text" name="DateRequested" size="12" value="<%=((String(Request.Form("DateRequested"))!="undefined")?Request.Form("DateRequested"):CurrentDate())%>" tabindex="10" maxlength="10" onChange="FormatDate(this);">
        <span style="font-size: 7pt">(mm/dd/yyyy)</span> </td>
    </tr>
    <tr> 
      <td nowrap>
        <input type="checkbox" name="IsReceived" tabindex="11" <%=((Request.Form("IsReceived")=="on")?"CHECKED":"")%> onClick="ChangeReceived();" class="chkstyle">
        Date Received:</td>
      <td nowrap> 
        <input type="text" name="DateReceived" size="12" value="<%=Request.Form("DateReceived")%>" tabindex="12" maxlength="10" onChange="FormatDate(this);">
        <span style="font-size: 7pt">(mm/dd/yyyy)</span> </td>
    </tr>
    <tr> 
      <td nowrap valign="top">Description:</td>
      <td nowrap valign="top">
        <textarea name="Description" rows="5" cols="65" tabindex="13" accesskey="L"><%=Request.Form("ServiceRequested")%></textarea>
      </td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" name="Save" value="Save" onClick="Save1();" tabindex="14" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="15" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action">
<input type="hidden" name="UserType">
<input type="hidden" name="UserID">
<input type="hidden" name="hInventoryID">
</form>
</body>
</html>
<%
rsRepairStatus.Close();
%>