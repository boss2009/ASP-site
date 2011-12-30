<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_action")) == "update") {
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))!="undefined")?String(Request.Form("WayBillNumber")).replace(/'/g, "''"):"");			
	var ScheduledArrivalDate = ((String(Request.Form("ScheduledArrivalDate"))=="undefined")?"1/1/1900":Request.Form("ScheduledArrivalDate"));		
	var DeliveryDateArranged = ((String(Request.Form("DeliveryDateArranged"))=="undefined")?"1/1/1900":Request.Form("DeliveryDateArranged"));			
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");	
	var MorningPickUp = null;
	if (String(Request.Form("Morning"))=="on") MorningPickUp = 1;
	if (String(Request.Form("Afternoon"))=="on") MorningPickUp = 0;
	var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
	rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
	rsShippingMethod.Source = "{call dbo.cp_eqpsrv_ship_method("+Request.Form("MM_recordId")+","+Request.Form("ShippingStatus")+","+Request.Form("ShippedBy")+","+Request.Form("ShippingMethod")+",'"+WayBillNumber+"','"+DeliveryDateArranged+"','"+ScheduledArrivalDate+"',"+MorningPickUp+",'"+Notes+"',"+Session("insStaff_id")+",0,'E',0)}";
	rsShippingMethod.CursorType = 0;
	rsShippingMethod.CursorLocation = 2;
	rsShippingMethod.LockType = 3;
	rsShippingMethod.Open();
	Response.Redirect("UpdateSuccessful2.asp?page=m009e0401.asp&intEquip_srv_id="+Request.Form("MM_recordId")+"&intShip_Dtl_id="+Request.Form("intShip_Dtl_id"));		
}

var rsEquipmentService = Server.CreateObject("ADODB.Recordset");
rsEquipmentService.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentService.Source = "{call dbo.cp_get_eqp_srv("+ Request.QueryString("intEquip_Srv_id") + ",0,0,'',1,'Q',0)}";
rsEquipmentService.CursorType = 0;
rsEquipmentService.CursorLocation = 2;
rsEquipmentService.LockType = 3;
rsEquipmentService.Open();

var rsUserType = Server.CreateObject("ADODB.Recordset");
rsUserType.ActiveConnection = MM_cnnASP02_STRING;

// + nOV.04.2005
//rsUserType.Source = "{call dbo.cp_FrmHdr_9A("+ Request.QueryString("intEquip_Srv_id") + ",0)}";
rsUserType.Source = "{call dbo.cp_FrmHdr_9("+ Request.QueryString("intEquip_Srv_id") + ",0)}";

rsUserType.CursorType = 0;
rsUserType.CursorLocation = 2;
rsUserType.LockType = 3;
rsUserType.Open();

var User = 0;
if (!rsUserType.EOF) {
	switch (String(rsUserType.Fields.Item("insEq_user_type").Value)) {
		//staff
		case "1":
			User = 5;
		break;
		//client
		case "3":
			User = 1;
		break;
		//school
		case "4":
			User = 2;
		break;
		//no user
		default:
			User = 3;
		break;
	} 
}

var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
rsShippingMethod.Source = "{call dbo.cp_eqpsrv_ship_method("+Request.QueryString("intEquip_srv_id")+",'',0,0,'','','',0,'',0,1,'Q',0)}";
rsShippingMethod.CursorType = 0;
rsShippingMethod.CursorLocation = 2;
rsShippingMethod.LockType = 3;
rsShippingMethod.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsShippingStatus = Server.CreateObject("ADODB.Recordset");
rsShippingStatus.ActiveConnection = MM_cnnASP02_STRING;
rsShippingStatus.Source = "{call dbo.cp_ship_rtn_status(0,'',0,'Q',0)}";
rsShippingStatus.CursorType = 0;
rsShippingStatus.CursorLocation = 2;
rsShippingStatus.LockType = 3;
rsShippingStatus.Open();

var rsMethod = Server.CreateObject("ADODB.Recordset");
rsMethod.ActiveConnection = MM_cnnASP02_STRING;
rsMethod.Source = "{call dbo.cp_shipping_method(0,0)}";
rsMethod.CursorType = 0;
rsMethod.CursorLocation = 2;
rsMethod.LockType = 3;
rsMethod.Open();

//Set bitIs_BackOrder = 1 for Equip. Service shipping
//Set bitIs_Backorder = 0 for Out Service
var intShip_Dtl_id = ((rsEquipmentService.Fields.Item("intShip_Dtl_id").Value!=null)?rsEquipmentService.Fields.Item("intShip_Dtl_id").Value:0);
	
var rsBoxes = Server.CreateObject("ADODB.Recordset");
rsBoxes.ActiveConnection = MM_cnnASP02_STRING;
rsBoxes.Source = "{call dbo.cp_eqpsrv_ship_box(0,"+intShip_Dtl_id+",0,0,"+Request.QueryString("intEquip_Srv_id")+",1,0,'Q',0)}";
rsBoxes.CursorType = 0;
rsBoxes.CursorLocation = 2;
rsBoxes.LockType = 3;
//Response.Redirect(rsBoxes.Source);
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
	<title>Shipping Method</title>
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
				document.frm0401.reset();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
		switch (document.frm0401.ShippingMethod.value) {
			//dynamex
			case "9":
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.Morning.disabled = false;
				document.frm0401.Afternoon.disabled = false;	
				document.frm0401.WayBillNumber.disabled = false;
			break;
			//picked up by client
			case "10":
				document.frm0401.DeliveryDateArranged.disabled = false;			
				document.frm0401.Morning.disabled = false;
				document.frm0401.Afternoon.disabled = false;
				document.frm0401.WayBillNumber.disabled = true;
			break;
			//taken by consultant
			case "1":
				document.frm0401.DeliveryDateArranged.disabled = false;			
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;
				document.frm0401.WayBillNumber.disabled = true;												
			break;
			//loomis
			case "4":
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;
				document.frm0401.WayBillNumber.disabled = false;
			break;
			//Third Party
			case "11":
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;				
				document.frm0401.WayBillNumber.disabled = false;
			break;			
			//none
			default:	
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;
				document.frm0401.WayBillNumber.disabled = true;
			break;
		}
	<%
	if (intShip_Dtl_id == 0) {
	%>
		document.frm0401.ShippingStatus.value = "<%=User%>";
		document.frm0401.ShippedBy.value = "<%=Session("insStaff_id")%>";
	<%
	}
	%>
		document.frm0401.ShippingStatus.focus();
	}

	function ChangeShippingMethod(){
		switch (document.frm0401.ShippingMethod.value) {
			//dynamex
			case "9":
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.DeliveryDateArranged.value="<%=CurrentDate()%>";				
				document.frm0401.ScheduledArrivalDate.value="<%=CurrentDate()%>";
				document.frm0401.Morning.disabled = false;
				document.frm0401.Afternoon.disabled = false;	
				document.frm0401.WayBillNumber.disabled = false;
			break;
			//picked up by client
			case "10":
				document.frm0401.DeliveryDateArranged.disabled = false;			
				document.frm0401.DeliveryDateArranged.value="<%=CurrentDate()%>";							
				document.frm0401.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0401.Morning.disabled = false;
				document.frm0401.Afternoon.disabled = false;
				document.frm0401.WayBillNumber.disabled = true;
			break;
			//taken by consultant
			case "1":
				document.frm0401.DeliveryDateArranged.disabled = false;			
				document.frm0401.DeliveryDateArranged.value = "";							
				document.frm0401.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;
				document.frm0401.WayBillNumber.disabled = true;												
			break;
			//loomis
			case "4":
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.DeliveryDateArranged.value = "";														
				document.frm0401.ScheduledArrivalDate.value=ForwardDay(1);
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;
				document.frm0401.WayBillNumber.disabled = false;
			break;
			//Third Party
			case "11":
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.DeliveryDateArranged.value = "";											
				document.frm0401.ScheduledArrivalDate.value="<%=CurrentDate()%>";													
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;				
				document.frm0401.WayBillNumber.disabled = false;
			break;			
			//none
			default:	
				document.frm0401.DeliveryDateArranged.disabled = false;
				document.frm0401.DeliveryDateArranged.value = "";																
				document.frm0401.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0401.Morning.disabled = true;
				document.frm0401.Afternoon.disabled = true;
				document.frm0401.WayBillNumber.disabled = true;
			break;
		}
	}

	function openWindow(page){
		if (page!='nothing') win1=window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function ListBoxes(){	
		openWindow('m009pop3.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>&intShip_Dtl_id=<%=intShip_Dtl_id%>');		
	}
	
	function Save(){
		if (!CheckTextArea(document.frm0401.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}	
		if (!CheckDate(document.frm0401.DeliveryDateArranged.value)){
			alert("Invalid Delivery Date Arranged.");
			document.frm0401.DeliveryDateArranged.focus();
			return ;
		}
		if (!CheckDate(document.frm0401.ScheduledArrivalDate.value)){
			alert("Invalid Scheduled Arrival Date.");
			document.frm0401.ScheduledArrivalDate.focus();
			return ;
		}
		document.frm0401.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0401">
<h5>Shipping Method</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Shipping Status:</td>
		<td nowrap><select name="ShippingStatus" tabindex="1" onChange="alert('Changing this field may affect integrity of inventory status.');" accesskey="F">
		<%
		while (!rsShippingStatus.EOF){
		%>
			<option value="<%=rsShippingStatus.Fields.Item("insRtn_to_User").Value%>" <%=((rsShippingMethod.Fields.Item("chrRtn_to_User").Value==String(rsShippingStatus.Fields.Item("insRtn_to_User").Value))?"SELECTED":"")%>><%=rsShippingStatus.Fields.Item("chvRtoUser_Desc").Value%>
		<%
			rsShippingStatus.MoveNext();
		}
		%>
		</select></td>
	</tr>  
	<tr> 
		<td nowrap>Shipped By:</td>
		<td nowrap><select name="ShippedBy" tabindex="2">
			<option value="0">(none)		
		<% 
		while (!rsStaff.EOF) {
		%>
			<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%=((rsStaff.Fields.Item("insStaff_id").Value==rsShippingMethod.Fields.Item("insShip_Staff_id").Value)?"SELECTED":"")%>><%=(rsStaff.Fields.Item("chvName").Value)%>
		<%
			rsStaff.MoveNext();
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Shipping Method:</td>
		<td nowrap><select name="ShippingMethod" onChange="ChangeShippingMethod();" tabindex="3">
			<option value="0">(none)
	<% 
	while (!rsMethod.EOF) {
		if (rsMethod.Fields.Item("bitis_active").Value == "1") {
	%>
			<option value="<%=(rsMethod.Fields.Item("intship_method_id").Value)%>" <%=((rsMethod.Fields.Item("intship_method_id").Value==rsShippingMethod.Fields.Item("insShip_Method_id").Value)?"SELECTED":"")%>><%=(rsMethod.Fields.Item("chvname").Value)%>
	<%
		}
		rsMethod.MoveNext();
	}
	%>
		</select></td>
    </tr>	
	<tr> 
		<td nowrap>Waybill Number:</td>
		<td nowrap><input type="text" name="WayBillNumber" value="<%=(rsShippingMethod.Fields.Item("chvWayBill_No").Value)%>" size="15" tabindex="4"></td>
	</tr>
	<tr>
		<td nowrap>Number of Boxes:</td>
		<td nowrap>
			<input type="text" name="NumberOfBoxes" size="2" maxlength="3" value="<%=count%>" tabindex="5" style="border: none" readOnly> Total Weight: 
			<input type="text" name="TotalWeight" size="4" value="<%=total%>" tabindex="6" style="border: none" readOnly> LB 
			<input type="button" value="Add/Update" tabindex="7" onClick="<%=((intShip_Dtl_id>0)?"ListBoxes();":"alert('Please save first, before adding shipping boxes.');")%>" class="btnstyle">
		</td>
	</tr>	
    <tr>
		<td nowrap>Delivery Date Arranged:</td>
		<td nowrap>
			<input type="text" name="DeliveryDateArranged" size="11" maxlength="10" value="<%=FilterDate(rsShippingMethod.Fields.Item("dtsDlvy_date").Value)%>" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr>
		<td nowrap>Scheduled Arrival Date:</td>
		<td nowrap>
			<input type="text" name="ScheduledArrivalDate" size="11" maxlength="10" value="<%=FilterDate(rsShippingMethod.Fields.Item("dtsSch_Arv_date").Value)%>" tabindex="9" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Picked Up:</td>
		<td nowrap>
			<input type="checkbox" name="Morning" tabindex="10" <%=((rsShippingMethod.Fields.Item("BitPkup_morning").Value=="1")?"CHECKED":"")%> class="chkstyle">Morning
			<input type="checkbox" name="Afternoon" tabindex="11" <%=((rsShippingMethod.Fields.Item("BitPkup_morning").Value=="0")?"CHECKED":"")%> class="chkstyle">Afternoon
		</td>
	</tr>
	<tr>
		<td nowrap valign="top">Shipping Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="12" accesskey="L"><%=rsShippingMethod.Fields.Item("chvNote_Desc").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="13" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="14" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="update">
<input type="hidden" name="MM_recordId" value="<%=rsShippingMethod.Fields.Item("intEquip_srv_id").Value %>">
<input type="hidden" name="intShip_Dtl_id" value="<%=intShip_Dtl_id%>">
</form>
</body>
</html>
<%
rsShippingMethod.Close();
rsStaff.Close();
rsMethod.Close();
%>