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
	var IsReturned = ((Request.Form("IsReturned")=="1")?"1":"0");	
	var DateReturned = ((String(Request.Form("DateReturned"))=="undefined")?"1/1/1900":Request.Form("DateReturned"));		
	var Description = String(Request.Form("Description")).replace(/'/g, "''");
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))=="undefined")?"":Request.Form("WayBillNumber"));
	var rsOutService = Server.CreateObject("ADODB.Recordset");
	rsOutService.ActiveConnection = MM_cnnASP02_STRING;
	rsOutService.Source = "{call dbo.cp_EqpSrv_Pfrm_SrvOut("+ Request.Form("MM_recordId") + "," + Request.Form("ShippedTo") + ",'" + Request.Form("DateShippedToVendor") +"',"+Request.Form("ShippingMethod") +","+ IsReturned+",'"+DateReturned+"','"+Request.Form("WayBillNumber")+"','"+Description+"',"+Session("insStaff_id")+",0,'E',0)}";
	rsOutService.CursorType = 0;
	rsOutService.CursorLocation = 2;
	rsOutService.LockType = 3;
	rsOutService.Open();
		
	//This is a Trigger to insert into user's services and notes page with I-Repair code
	//Another Trigger to update repair status to Repair Completed
	//Added another Trigger here to change that inventory status back to its original status before the repair
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

		var CheckEquipmentService = Server.CreateObject("ADODB.Command");
		CheckEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
		CheckEquipmentService.CommandText = "dbo.cp_chk_eqpsrv_type";
		CheckEquipmentService.CommandType = 4;
		CheckEquipmentService.CommandTimeout = 0;
		CheckEquipmentService.Prepared = true;
		CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("RETURN_VALUE", 3, 4));
		CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("@intEquip_srv_id", 3, 1,10000,Request.QueryString("intEquip_srv_id")));
		CheckEquipmentService.Parameters.Append(CheckEquipmentService.CreateParameter("@insRtnFlag", 2, 2));
		CheckEquipmentService.Execute();
	
		var TransactionType = "";
		if (CheckEquipmentService.Parameters.Item("@insRtnFlag").Value > 3) {
			TransactionType = "Buyout";
		} else {
			TransactionType = "Loan";
		}		
	
		Description = Trim(rsUserType.Fields.Item("chvInventory_Name").Value) + "\n" + Description;
			
		var Year = CurrentYear();
		var Cycle = CurrentMonth();
		//User Type
		switch (String(rsUserType.Fields.Item("insEq_user_type").Value)) {
			case "3":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				if (TransactionType=="Buyout") {
					rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+DateReturned+"',"+Year+","+Cycle+",0,'"+Description.replace(/'/g, "''")+"','2202000000000000000000000000000000000000',0,'A',0)}";
				} else {
					rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+DateReturned+"',"+Year+","+Cycle+",0,'"+Description.replace(/'/g, "''")+"','2205000000000000000000000000000000000000',0,'A',0)}";
				}
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();
			break;
			case "4":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+rsUserType.Fields.Item("intEq_user_id").Value+",0,'"+DateReturned+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Description.replace(/'/g, "''")+"','2206000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();			
			break;
		}
	}
		
	Response.Redirect("UpdateSuccessful.asp?page=m009e0302.asp&intEquip_Srv_id="+Request.QueryString("intEquip_Srv_id"))	
}

var rsEquipmentService = Server.CreateObject("ADODB.Recordset");
rsEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentService.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",0,0,'',1,'Q',0)}";
rsEquipmentService.CursorType = 0;
rsEquipmentService.CursorLocation = 2;
rsEquipmentService.LockType = 3;
rsEquipmentService.Open();

var rsOutService = Server.CreateObject("ADODB.Recordset");
rsOutService.ActiveConnection = MM_cnnASP02_STRING;
rsOutService.Source = "{call dbo.cp_EqpSrv_Pfrm_SrvOut("+ Request.QueryString("intEquip_srv_id") + ",0,'',0,0,'','','',"+Session("insStaff_id")+",1,'Q',0)}"
rsOutService.CursorType = 0;
rsOutService.CursorLocation = 2;
rsOutService.LockType = 3;
rsOutService.Open();

var IsNew = ((rsOutService.EOF)?true:false);

var rsVendor = Server.CreateObject("ADODB.Recordset");
rsVendor.ActiveConnection = MM_cnnASP02_STRING;
rsVendor.Source = "{call dbo.cp_company2(0,'',0,0,0,0,0,1,0,'',0,'Q',0)}"
rsVendor.CursorType = 0;
rsVendor.CursorLocation = 2;
rsVendor.LockType = 3;
rsVendor.Open();

var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
rsShippingMethod.Source = "{call dbo.cp_shipping_method(0,0)}";
rsShippingMethod.CursorType = 0;
rsShippingMethod.CursorLocation = 2;
rsShippingMethod.LockType = 3;
rsShippingMethod.Open();

var intSrv_dtl_id = 0;
if (!IsNew) {
	intSrv_dtl_id = ((rsOutService.Fields.Item("intShip_dtl_id").Value != null)?rsOutService.Fields.Item("intShip_dtl_id").Value:0);
}
	
var rsBoxes = Server.CreateObject("ADODB.Recordset");
rsBoxes.ActiveConnection = MM_cnnASP02_STRING;
rsBoxes.Source = "{call dbo.cp_eqpsrv_ship_box(0,"+intSrv_dtl_id+",0,0,"+Request.QueryString("intEquip_Srv_id")+",0,0,'Q',0)}";
rsBoxes.CursorType = 0;
rsBoxes.CursorLocation = 2;
rsBoxes.LockType = 3;
rsBoxes.Open();
var count = 0;
var total = 0;
while (!rsBoxes.EOF) {
	count++;
	total = total + rsBoxes.Fields.Item("insBox_Wgt").Value;
	rsBoxes.MoveNext();
}
%>
<html>
<head>
	<title>Out Service</title>
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
				document.frm0302.reset();
			break;
			case 76 :
				//alert("L");
				window.location.href='m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>';
			break;
		}
	}
	</script>		
	<script language="Javascript">
	function ListBoxes(){	
		openWindow('m009pop2.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intSrv_dtl_id=<%=((!IsNew)?rsOutService.Fields.Item("intShip_dtl_id").Value:"0")%>');
	}
	
	function ChangeShippingMethod(){
		switch (document.frm0302.ShippingMethod.value) {
			//dynamex
			case "9":
				document.frm0302.WayBillNumber.disabled = false;
			break;
			//Loomis
			case "4":
				document.frm0302.WayBillNumber.disabled = false;
			break;
			//Third Party
			case "11":
				document.frm0302.WayBillNumber.disabled = false;
			break;
			default :
				document.frm0302.WayBillNumber.disabled = true;
			break;
		}
	}
	
	function ChangeReturned() {
		if (document.frm0302.IsReturned.checked) {
			document.frm0302.DateReturned.disabled = false;
			if (Trim(document.frm0302.DateReturned.value)=="") {			
				document.frm0302.DateReturned.value = "<%=CurrentDate()%>";
			}
		} else {
			document.frm0302.DateReturned.disabled = true;		
			document.frm0302.DateReturned.value = "";			
		}		
	}

	function PrintShippingLabel(){
		document.frm0302.action = "m009e0404.asp?intEquip_srv_id=<%=Request.QueryString("intEquip_srv_id")%>";
		document.frm0302.target = "_blank";
		document.frm0302.submit();	
	}
		
	function Init() {
		if (!document.frm0302.IsReturned.checked) {
			document.frm0302.DateReturned.disabled = true;		
			document.frm0302.DateReturned.value = "";			
		}		
		document.frm0302.IsReturned.focus();
	}	

	function openWindow(page){
		if (page!='nothing') win1 = window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function Save(){
		if (!CheckDate(document.frm0302.DateShippedToVendor.value)) {
			alert("Invalid Date Shipped.");
			document.frm0302.DateShippedToVendor.focus();
			return ;
		}
		if (document.frm0302.IsReturned.checked) {
			if (!CheckDate(document.frm0302.DateReturned.value)) {
				alert("Invalid Date Returned.");
				document.frm0302.DateReturned.focus();
				return ;
			}
		} else {
			document.frm0302.InsertNote.value = "False";
		}
		if (!CheckTextArea(document.frm0302.Description, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}		
		document.frm0302.submit();
	}
	</script>	
</head>
<body onLoad="Init();">
<form name="frm0302" method="POST" action="<%=MM_editAction%>">
<h5>Out Service</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Shipped To:</td>
		<td nowrap colspan="3"><select name="ShippedTo" tabindex="1" accesskey="F">
		<%
		var vendor = ((IsNew)?rsEquipmentService.Fields.Item("insVendor_id").Value:rsOutService.Fields.Item("intVendor_id").Value);
		while (!rsVendor.EOF) {			
		%>
			<option value="<%=(rsVendor.Fields.Item("intCompany_id").Value)%>" <%=((rsVendor.Fields.Item("intCompany_id").Value==vendor)?"SELECTED":"")%>><%=rsVendor.Fields.Item("chvCompany_Name").Value%>
		<%
			rsVendor.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td valign="top" nowrap>Address:</td>
		<td valign="top" colspan="3">
			<%=((!IsNew)?rsOutService.Fields.Item("chvAddress").Value:"")%>&nbsp;<br>
			<%=((!IsNew)?rsOutService.Fields.Item("chvCity").Value:"")%>&nbsp;<%=((!IsNew)?rsOutService.Fields.Item("chrprvst_abbv").Value:"")%><br>			
			<%=((!IsNew)?rsOutService.Fields.Item("chvcntry_name").Value:"")%>&nbsp;<%=((!IsNew)?rsOutService.Fields.Item("chvPostal_zip").Value:"")%><br><br>			
		</td>
	</tr>
	<tr>
		<td nowrap>Date Shipped:</td>
		<td nowrap>
			<input type="text" name="DateShippedToVendor" value="<%=((!IsNew)?FilterDate(rsOutService.Fields.Item("dtsDate_Ship_to_Vendor").Value):"")%>" size="11" maxlength="10" tabindex="2" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Shipping Method:</td>
		<td nowrap><select name="ShippingMethod" onChange="ChangeShippingMethod();" tabindex="5" style="width: 150px">	
			<option value="0">N/A		
		<%
		while (!rsShippingMethod.EOF) {
			if (rsShippingMethod.Fields.Item("bitis_active").Value == "1") {
		%>
			<option value="<%=rsShippingMethod.Fields.Item("intship_method_id").Value%>" <%if (!IsNew) Response.Write((rsShippingMethod.Fields.Item("intship_method_id").Value==rsOutService.Fields.Item("insShip_Method_id").Value)?"SELECTED":"")%>><%=rsShippingMethod.Fields.Item("chvname").Value%>
		<%
			}
			rsShippingMethod.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td nowrap><input type="checkbox" name="IsReturned" value="1" tabindex="3" accesskey="F" <%if (!IsNew) Response.Write((rsOutService.Fields.Item("bitIs_Reurn_to_ASP").Value=="1")?"CHECKED":"")%> onClick="ChangeReturned();" class="chkstyle">Returned to AT-BC:</td>
		<td nowrap>
			<input type="text" name="DateReturned" value="<%=((!IsNew)?FilterDate(rsOutService.Fields.Item("dtsReurn_to_ASP").Value):"")%>" size="11" maxlength="10" tabindex="4" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Waybill Number:</td>
		<td nowrap><input type="text" name="WayBillNumber" value="<%=((!IsNew)?rsOutService.Fields.Item("chvWayBill_No").Value:"")%>" size="15" tabindex="5"></td>
	</tr>
    <tr>
		<td nowrap>Number of Boxes:</td>
		<td nowrap colspan="2">
			<input type="text" name="NumberOfBoxes" value="<%=count%>" size="2" maxlength="3" tabindex="6" style="text-align: right;border: none" readonly>&nbsp;&nbsp;&nbsp;Total Weight:
			<input type="text" name="TotalWeight" value="<%=total%>" size="7" tabindex="7" style="text-align: right; border: none" readonly>&nbsp;&nbsp;LB
			<input type="button" value="Add/Update" tabindex="8" onClick="<%=((!IsNew)?"ListBoxes();":"alert('Please save first, before adding shipping boxes.');")%>" class="btnstyle">
		</td>
    </tr>
	<tr>
		<td valign="top">Description:</td>
		<td valign="top" colspan="3"><textarea name="Description" cols="65" rows="5" tabindex="9" accesskey="L"><%=((!IsNew)?rsOutService.Fields.Item("chvPfrm_Dscr_Out").Value:"")%></textarea>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="10" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Print Shipping Label" tabindex="11" onClick="PrintShippingLabel();" class="btnstyle"></td>		
		<td><input type="button" value="Close" tabindex="12" onClick="window.location.href='m009FS01.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>'" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=Request.QueryString("intEquip_Srv_id")%>">
<input type="hidden" name="StreetAddress" value="<%=((!IsNew)?rsOutService.Fields.Item("chvAddress").Value:"")%>">
<input type="hidden" name="City" value="<%=((!IsNew)?rsOutService.Fields.Item("chvCity").Value:"")%>">
<input type="hidden" name="UserName" value="<%=((!IsNew)?rsOutService.Fields.Item("chvShip_to").Value:"")%>">
<input type="hidden" name="ProvinceState" value="<%=((!IsNew)?rsOutService.Fields.Item("chrprvst_abbv").Value:"")%>">
<input type="hidden" name="PostalCode" value="<%=((!IsNew)?FormatPostalCode(rsOutService.Fields.Item("chvPostal_zip").Value):"")%>">
<input type="hidden" name="InsertNote" value="<%if (!IsNew) { Response.Write((rsOutService.Fields.Item("bitIs_Reurn_to_ASP").Value=="1")?"False":"True") } else {Response.Write("True")}%>">
</form>
</body>
</html>
<%
rsOutService.Close();
%>