<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update"))=="true"){
	var IsCompleted = ((Request.Form("IsCompleted")=="1")?"1":"0");	
	var DateCompleted = ((String(Request.Form("DateCompleted"))=="undefined")?"1/1/1900":Request.Form("DateCompleted"));		
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var LabourCost = ((String(Request.Form("LabourCost"))=="undefined")?"0":Request.Form("LabourCost"));
	var LabourCostOption = ((String(Request.Form("LabourCostOption"))=="undefined")?null:Request.Form("LabourCostOption"));
	var rsInService = Server.CreateObject("ADODB.Recordset");
	rsInService.ActiveConnection = MM_cnnASP02_STRING;
	rsInService.Source = "{call dbo.cp_EqpSrv_Perform("+ Request.Form("MM_recordId") + ","+IsCompleted+",'"+DateCompleted+"',"+Request.Form("RepairedBy")+",'"+Request.Form("ReasonForRepair")+"',"+Request.Form("TypeOfRepair")+","+LabourCost+","+LabourCostOption+","+Request.Form("PartsCost")+","+Request.Form("ServiceHours")+","+Request.Form("ServiceMinutes")+",'"+Description+"',0,'E',0)}";
	rsInService.CursorType = 0;
	rsInService.CursorLocation = 2;
	rsInService.LockType = 3;
//	Response.Redirect(rsInService.Source);
	rsInService.Open();

	//This is a Trigger to insert into user's services and notes page with I-Repair code
	//Another Trigger for updating repair status to Repair Complete
	if (String(Request.Form("InsertNote"))=="True") {
		var rsInventoryOldStatus = Server.CreateObject("ADODB.Recordset");
		rsInventoryOldStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryOldStatus.Source = "select intEquip_Srv_id, intEquip_Set_id, insOldEquip_status_id from tbl_eqp_srv where intEquip_Srv_id = " + Request.QueryString("intEquip_Srv_id");
		rsInventoryOldStatus.CursorType = 0;
		rsInventoryOldStatus.CursorLocation = 2;
		rsInventoryOldStatus.LockType = 3;
		rsInventoryOldStatus.Open();
		
		var rsSetInventoryStatus = Server.CreateObject("ADODB.Recordset");
		rsSetInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetInventoryStatus.Source = "{call dbo.cp_update_eqpivtry_status("+rsInventoryOldStatus.Fields.Item("intEquip_Set_id").Value+","+rsInventoryOldStatus.Fields.Item("insOldEquip_status_id").Value+",0)}";
		rsSetInventoryStatus.CursorType = 0;
		rsSetInventoryStatus.CursorLocation = 2;
		rsSetInventoryStatus.LockType = 3;
		rsSetInventoryStatus.Open();
	
		var rsEquipmentRepairStatus = Server.CreateObject("ADODB.Recordset");
		rsEquipmentRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsEquipmentRepairStatus.Source = "{call dbo.cp_eqpsrv_repsts("+Request.QueryString("intEquip_Srv_id")+",2,0,'E',0)}";
		rsEquipmentRepairStatus.CursorType = 0;
		rsEquipmentRepairStatus.CursorLocation = 2;
		rsEquipmentRepairStatus.LockType = 3;
		rsEquipmentRepairStatus.Open();
	
		var rsUserType = Server.CreateObject("ADODB.Recordset");
		rsUserType.ActiveConnection = MM_cnnASP02_STRING;
		rsUserType.Source = "{call dbo.cp_FrmHdr_9A("+ Request.QueryString("intEquip_Srv_id") + ",0)}";
		rsUserType.CursorType = 0;
		rsUserType.CursorLocation = 2;
		rsUserType.LockType = 3;
		rsUserType.Open();

		//Obsolete sp.
//		var CheckEquipmentService = Server.CreateObject("ADODB.Command");
//		CheckEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
//		CheckEquipmentService.CommandText = "dbo.cp_chk_eqpsrv_type";
//		CheckEquipmentService.CommandType = 4;
//		CheckEquipmentService.CommandTimeout = 0;
//		CheckEquipmentService.Prepared = true;
//		CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("RETURN_VALUE", 3, 4));
//		CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("@intEquip_srv_id", 3, 1,10000,Request.QueryString("intEquip_srv_id")));
//		CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("@insRtnFlag", 2, 2));
//		CheckEquipmentService.Execute();
//		if (CheckEquipmentService.Parameters.Item("@insRtnFlag").Value > 3) {
//			TransactionType = "Buyout";
//		} else {
//			TransactionType = "Loan";
//		}
	
		var TransactionType = "";
		
		switch (String(rsInventoryOldStatus.Fields.Item("insOldEquip_status_id").Value)) {
			//Loaned
			case "3":
				TransactionType = "Loan";
			break;
			//Buyout			
			case "11":
				TransactionType = "Sold";
			break;
			//Buyout			
			case "14":
				TransactionType = "Sold";			
			break;
			//Buyout			
			case "15":
				TransactionType = "Sold";			
			break;
			//Buyout			
			case "16":
				TransactionType = "Sold";			
			break;
			//Buyout			
			case "17":
				TransactionType = "Sold";			
			break;
		}
		
		Description = Trim(rsUserType.Fields.Item("chvInventory_Name").Value) + "\n" + Description;
		
		var Year = CurrentYear();
		var Cycle = CurrentMonth();
		//User Type
		switch (String(rsUserType.Fields.Item("insEq_user_type").Value)) {
			case "3":
				//Only insert service request if the item was on loan or sold.
				if (TransactionType != "") {
					var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
					rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
					if (TransactionType == "Buyout") {
						rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+DateCompleted+"',"+Year+","+Cycle+",0,'"+Description.replace(/'/g, "''")+"','3500000000000000000000000000000000000000',0,'A',0)}";
					} else {
						rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+DateCompleted+"',"+Year+","+Cycle+",0,'"+Description.replace(/'/g, "''")+"','3600000000000000000000000000000000000000',0,'A',0)}";				
					}					
					rsServiceRequested.CursorType = 0;
					rsServiceRequested.CursorLocation = 2;
					rsServiceRequested.LockType = 3;
					rsServiceRequested.Open();
				}
			break;
			case "4":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+DateCompleted+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Description.replace(/'/g, "''")+"','3700000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();			
			break;
		}
	}
		
	Response.Redirect("UpdateSuccessful.asp?page=m009e0301.asp&intEquip_Srv_id="+Request.QueryString("intEquip_Srv_id"))	
}

var rsInService = Server.CreateObject("ADODB.Recordset");
rsInService.ActiveConnection = MM_cnnASP02_STRING;
rsInService.Source = "{call dbo.cp_EqpSrv_Perform("+ Request.QueryString("intEquip_srv_id") + ",0,'',0,'',0,0.0,0,0.0,0,0,'',1,'Q',0)}"
rsInService.CursorType = 0;
rsInService.CursorLocation = 2;
rsInService.LockType = 3;
rsInService.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsRepairReason = Server.CreateObject("ADODB.Recordset");
rsRepairReason.ActiveConnection = MM_cnnASP02_STRING;
rsRepairReason.Source = "{call dbo.cp_eqsrv_repair_reason(0,'',0,0,0,'Q',0)}"
rsRepairReason.CursorType = 0;
rsRepairReason.CursorLocation = 2;
rsRepairReason.LockType = 3;
rsRepairReason.Open();
%>
<html>
<head>
	<title>In Service</title>
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
			case 85:
				//alert("U");
				document.frm0301.reset();
			break;
			case 76 :
				//alert("L");
				window.location.href='m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>';
			break;
		}
	}
	</script>		
	<script language="Javascript">
	function ChangeReason(){
		if (document.frm0301.TypeOfRepair.value=="0"){
			switch (String(document.frm0301.ReasonForRepair.value)) {
		<%
		while (!rsRepairReason.EOF) {
			if (rsRepairReason.Fields.Item("bitIs_Labor_cost").Value==1) {
		%>			
				case "<%=rsRepairReason.Fields.Item("insRepair_Reason_Id").Value%>":
					document.frm0301.LabourCost.disabled = false;
					document.frm0301.LabourCostOption.disabled = false;							
				break;
		<%
			}
			rsRepairReason.MoveNext();
		}
		%>				
			}
		} else {
			document.frm0301.LabourCost.disabled = true;
			document.frm0301.LabourCostOption.disabled = true;		
		}
		CalculateLabour();
	}
	
	function ChangeCompleted() {
		if (document.frm0301.IsCompleted.checked) {
			document.frm0301.DateCompleted.disabled = false;
			if (Trim(document.frm0301.DateCompleted.value)=="") {
				document.frm0301.DateCompleted.value = "<%=CurrentDate()%>";
			}
		} else {
			document.frm0301.DateCompleted.disabled = true;
			document.frm0301.DateCompleted.value = "";								
		}		
	}
	
	function Init() {
		ChangeReason();
		if (!document.frm0301.IsCompleted.checked) {
			document.frm0301.DateCompleted.disabled = true;
			document.frm0301.DateCompleted.value = "";								
		}		

		document.frm0301.IsCompleted.focus();
	}	
	
	function Save(){
		if (document.frm0301.IsCompleted.checked) {
			if (!CheckDate(document.frm0301.DateCompleted.value)) {
				alert("Invalid Date Completed.");
				document.frm0301.DateCompleted.focus();
				return ; 
			}
		} else {
			document.frm0301.InsertNote.value="False";
		}
		if (!document.frm0301.LabourCost.disabled) {
			if (isNaN(document.frm0301.LabourCost.value)) {
				alert("Invalid Labour Cost.");
				document.frm0301.LabourCost.focus();
				return ; 
			}
		}
		if (isNaN(document.frm0301.PartsCost.value)){
			alert("Invalid Parts Cost.");
			document.frm0301.PartsCost.focus();
			return ;			
		}
		if (Trim(document.frm0301.LabourCost.value)=="") document.frm0301.LabourCost.value = 0;		
		if (Trim(document.frm0301.PartsCost.value)=="") document.frm0301.PartsCost.value = 0;
		if (!CheckTextArea(document.frm0301.Description, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}		
		document.frm0301.submit();
	}
	
	function CalculateLabour(){
		document.frm0301.LabourCost.value = document.frm0301.ServiceHours.value * 50 + (document.frm0301.ServiceMinutes.value / 60) * 50;
		CalculateTotal();
	}
	
	function CalculateTotal(){
		if (!document.frm0301.LabourCost.disabled) {
			document.frm0301.TotalCost.value = Number(document.frm0301.LabourCost.value) + Number(document.frm0301.PartsCost.value);
		} else {
			document.frm0301.TotalCost.value = Number(document.frm0301.PartsCost.value);		
		}
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0301" method="POST" action="<%=MM_editAction%>">
<h5>In Service</h5>
<table cellpadding="1" cellspacing="1">
	<tr>		
		<td nowrap><input type="checkbox" name="IsCompleted" value="1" tabindex="1" accesskey="F" <%=((rsInService.Fields.Item("bitIs_Completed").Value=="1")?"CHECKED":"")%> onClick="ChangeCompleted();" class="chkstyle">Date Completed:</td>
		<td nowrap>
			<input type="text" name="DateCompleted" value="<%=FilterDate(rsInService.Fields.Item("dtsCompleted_Date").Value)%>" size="11" maxlength="10" tabindex="2" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Repaired By:</td>
		<td nowrap><select name="RepairedBy" tabindex="3" style="width: 150px">	
			<option value="0">(n/a)		
		<%
		var staff = ((rsInService.Fields.Item("insRepair_Staff_id").Value>0)?rsInService.Fields.Item("insRepair_Staff_id").Value:Session("insStaff_id"));
		while (!rsStaff.EOF) {
		%>
			<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==staff)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvname").Value%>
		<%
			rsStaff.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Reason For Repair:</td>
		<td nowrap><select name="ReasonForRepair" tabindex="4" style="width: 150px" onChange="ChangeReason();">
		<%
		rsRepairReason.ReQuery();
		while (!rsRepairReason.EOF) {
		%>
			<option value="<%=rsRepairReason.Fields.Item("insRepair_Reason_Id").Value%>" <%=((String(rsRepairReason.Fields.Item("insRepair_Reason_Id").Value)==rsInService.Fields.Item("chrReason_Repair").Value)?"SELECTED":"")%>><%=rsRepairReason.Fields.Item("chvRepair_Reason_Desc").Value%>
		<%
		rsRepairReason.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Type of Repair:</td>
		<td nowrap><select name="TypeOfRepair" tabindex="5" style="width: 150px" onChange="ChangeReason();">
			<option value="null">N/A
			<option value="1" <%=((rsInService.Fields.Item("bitIs_Covered_Warnty").Value=="1")?"SELECTED":"")%>>Covered by warranty
			<option value="0" <%=((rsInService.Fields.Item("bitIs_Covered_Warnty").Value=="0")?"SELECTED":"")%>>Not covered by warranty
		</select></td>
	</tr>
	<tr>
		<td nowrap>Service:</td>
		<td nowrap>
			<select name="ServiceHours" tabindex="6" onChange="CalculateLabour();">
				<option value="0" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="0")?"SELECTED":"")%>>0
				<option value="1" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="1")?"SELECTED":"")%>>1
				<option value="2" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="2")?"SELECTED":"")%>>2
				<option value="3" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="3")?"SELECTED":"")%>>3
				<option value="4" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="4")?"SELECTED":"")%>>4
				<option value="5" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="5")?"SELECTED":"")%>>5
				<option value="6" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="6")?"SELECTED":"")%>>6
				<option value="7" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="7")?"SELECTED":"")%>>7
				<option value="8" <%=((rsInService.Fields.Item("insSrv_hrs").Value=="8")?"SELECTED":"")%>>8
			</select>
			Hours
			<select name="ServiceMinutes" tabindex="7" onChange="CalculateLabour();">
				<option value="0" <%=((rsInService.Fields.Item("insSrv_minutes").Value=="0")?"SELECTED":"")%>>0
				<option value="15" <%=((rsInService.Fields.Item("insSrv_minutes").Value=="15")?"SELECTED":"")%>>15
				<option value="30" <%=((rsInService.Fields.Item("insSrv_minutes").Value=="30")?"SELECTED":"")%>>30
				<option value="45" <%=((rsInService.Fields.Item("insSrv_minutes").Value=="45")?"SELECTED":"")%>>45
			</select>
			Minutes
		</td>
	</tr>	
	<tr>
		<td nowrap>Labour Cost:</td>
		<td nowrap>
			$<input type="text" name="LabourCost" value="<%=rsInService.Fields.Item("fltLabour_Cost").Value%>" size="10" onKeypress="AllowNumericOnly();" onChange="CalculateTotal();" tabindex="8">
			<select name="LabourCostOption" tabindex="9">
				<option value="null">N/A
				<option value="1" <%=((rsInService.Fields.Item("bitLC_Status").Value=="1")?"SELECTED":"")%>>Charge
				<option value="0" <%=((rsInService.Fields.Item("bitLC_Status").Value=="0")?"SELECTED":"")%>>Waive service charge
			</select>
		</td>
	</tr>
	<tr>
		<td nowrap>Parts Cost:</td>
		<td nowrap>$<input type="text" name="PartsCost" value="<%=rsInService.Fields.Item("fltParts_Cost").Value%>" size="10" onKeypress="AllowNumericOnly();" onChange="CalculateTotal();" tabindex="10"></td>
	</tr>
	<tr>
		<td nowrap>Total Cost:</td>
		<td nowrap>$<input type="text" name="TotalCost" style="border: none" readonly tabindex="11" size="10" value="0"></td>
	</tr>
	<tr>
		<td valign="top">Description:</td>
		<td valign="top"><textarea name="Description" cols="65" rows="5" tabindex="12" accesskey="L"><%=rsInService.Fields.Item("chvNote_Desc").Value%></textarea>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="12" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="13" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="14" onClick="window.location.href='m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>'" class="btnstyle"></td>		
	</tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=Request.QueryString("intEquip_Srv_id")%>">
<input type="hidden" name="InsertNote" value="<%=((rsInService.Fields.Item("bitIs_Completed").Value=="1")?"False":"True")%>">
</form>
</body>
</html>
<%
rsInService.Close();
%>