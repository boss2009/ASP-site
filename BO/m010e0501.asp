<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var intShip_dtl_id = 0;
if (!rsBuyout.EOF) {
	if (rsBuyout.Fields.Item("intShip_dtl_id").Value != null) intShip_dtl_id = rsBuyout.Fields.Item("intShip_dtl_id").Value;
} 

var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_buyout_ship_notes("+ intShip_dtl_id + ",0,'',"+Session("insStaff_id")+",0,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();

var intShip_notes_id = 0;
if (!rsNotes.EOF) {
	if (rsNotes.Fields.Item("intShip_notes_id").Value != null) intShip_notes_id = rsNotes.Fields.Item("intShip_notes_id").Value;
}

if (String(Request("MM_action")) == "update") {
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))!="undefined")?String(Request.Form("WayBillNumber")).replace(/'/g, "''"):"");			
	var MorningPickedUp = null;
	if (String(Request.Form("PickedUp"))!="undefined") MorningPickedUp = Request.Form("PickedUp");
	var DeliveryDate = "1/1/1900";
	if (String(Request.Form("DeliveryDate"))!="undefined") DeliveryDate = Request.Form("DeliveryDate");	
	var IsProcessed = ((String(Request.Form("IsProcessed"))=="on")?"1":"0");	
	var IsDelivered = ((String(Request.Form("IsDelivered"))=="on")?"1":"0");	
	var rsMethod = Server.CreateObject("ADODB.Recordset");
	rsMethod.ActiveConnection = MM_cnnASP02_STRING;
	rsMethod.Source = "{call dbo.cp_buyout_ship_method3("+Request.Form("MM_recordId")+","+Request.QueryString("intBuyout_req_id")+","+IsProcessed+",'"+Request.Form("DateProcessed")+"',"+Request.Form("ProcessedBy")+","+Request.Form("ShippingMethod")+","+intShip_notes_id+",'"+WayBillNumber+"',"+Request.Form("NumberOfBoxes")+","+IsDelivered+",'"+DeliveryDate+"','"+Request.Form("ScheduledArrivalDate")+"',"+MorningPickedUp+",0,'E',0)}";
	rsMethod.CursorType = 0;
	rsMethod.CursorLocation = 2;
	rsMethod.LockType = 3;
	rsMethod.Open();

	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	rsNotes.Close();
	if (intShip_notes_id != 0) {
		rsNotes.Source = "{call dbo.cp_buyout_ship_notes(0,"+intShip_notes_id+",'"+Notes+"',"+Session("insStaff_id")+",0,'E',0)}";	
	} else {
		rsNotes.Source = "{call dbo.cp_buyout_ship_notes("+intShip_dtl_id+","+intShip_notes_id+",'"+Notes+"',"+Session("insStaff_id")+",0,'A',0)}";	
	}
	rsNotes.Open();	

	//Trigger for changing Buyout Status to Buyout Processed
	//Trigger to change Buyout Process to Ready To Invoice	
	if (String(Request.Form("Processed"))=="True") {
		var rsSetBuyoutStatus = Server.CreateObject("ADODB.Recordset");
		rsSetBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetBuyoutStatus.Source = "{call dbo.cp_update_buyout_status("+Request.QueryString("intBuyout_req_id")+",4,0)}";
		rsSetBuyoutStatus.CursorType = 0;
		rsSetBuyoutStatus.CursorLocation = 2;
		rsSetBuyoutStatus.LockType = 3;
		rsSetBuyoutStatus.Open();

		var rsSetBuyoutProcess = Server.CreateObject("ADODB.Recordset");
		rsSetBuyoutProcess.ActiveConnection = MM_cnnASP02_STRING;
		rsSetBuyoutProcess.Source = "update tbl_buyout_request set insBuyout_Prc_id = 4 where intBuyout_req_id = " + Request.QueryString("intBuyout_req_id");
		rsSetBuyoutProcess.CursorType = 0;
		rsSetBuyoutProcess.CursorLocation = 2;
		rsSetBuyoutProcess.LockType = 3;
		rsSetBuyoutProcess.Open();
	}
		
	//Trigger to change Buyout Status to Buyout Delivered
	//Trigger to insert into client/institution services and notes page with I-Tech code
	if (String(Request.Form("Delivered"))=="True") {
		var rsSetBuyoutStatus = Server.CreateObject("ADODB.Recordset");
		rsSetBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetBuyoutStatus.Source = "{call dbo.cp_update_buyout_status("+Request.QueryString("intBuyout_req_id")+",3,0)}";
		rsSetBuyoutStatus.CursorType = 0;
		rsSetBuyoutStatus.CursorLocation = 2;
		rsSetBuyoutStatus.LockType = 3;
		rsSetBuyoutStatus.Open();				
	
		Notes = "Buyout Request: " + Request.QueryString("intBuyout_req_id") + " shipped.\n" + Notes;	
		var Year = CurrentYear();
		var Cycle = CurrentMonth();
		//User Type
		switch (String(rsBuyout.Fields.Item("insEq_user_type").Value)) {
			//Client
			case "3":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsBuyout.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','3200000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();
			break;
			//Institution
			case "4":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+rsBuyout.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','2F00000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();			
			break;
		}
	}		
	Response.Redirect("m010e0501.asp?intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));
}

if (String(Request("MM_action")) == "insert") {
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))!="undefined")?String(Request.Form("WayBillNumber")).replace(/'/g, "''"):"");			
	var MorningPickedUp = null;
	if (String(Request.Form("PickedUp"))!="undefined") MorningPickedUp = Request.Form("PickedUp");
	var IsProcessed = ((String(Request.Form("IsProcessed"))=="on")?"1":"0");	
	var IsDelivered = ((String(Request.Form("IsDelivered"))=="on")?"1":"0");	
	var cmdInsertShipDetail = Server.CreateObject("ADODB.Command");
	cmdInsertShipDetail.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertShipDetail.CommandText = "dbo.cp_Buyout_Ship_Method3";
	cmdInsertShipDetail.CommandType = 4;
	cmdInsertShipDetail.CommandTimeout = 0;
	cmdInsertShipDetail.Prepared = true;
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intRecID", 3, 1,1,0));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intBuyout_Req_id", 3, 1,1,Request.QueryString("intBuyout_req_id")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@bitIs_Process", 2, 1,1,IsProcessed));	
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsUser_Ship_date", 200, 1,30,Request.Form("DateProcessed")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insShip_Staff_id", 2, 1,1,Request.Form("ProcessedBy")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insShip_Method_id", 2, 1,1,Request.Form("ShippingMethod")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intShip_notes_id", 3, 1,1,intShip_notes_id));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@chvWayBill_No", 200, 1,20,WayBillNumber));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insNum_of_Boxes", 2, 1,1,Request.Form("NumberOfBoxes")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@bitIs_Dlvy", 2, 1,1,IsDelivered));		
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsDlvy_date", 200, 1,30,Request.Form("DeliveryDate")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsSch_Arv_date", 200, 1,30,Request.Form("ScheduledArrivalDate")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@BitPkup_morning", 2, 1,1,MorningPickedUp));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intRtnValue", 3, 2));
	cmdInsertShipDetail.Execute();

	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_buyout_ship_notes("+ cmdInsertShipDetail.Parameters.Item("@intRtnValue").Value + ",0,'"+Notes+"',"+Session("insStaff_id")+",0,'A',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	
	//Trigger for changing Buyout Status to Buyout Processed
	//Trigger to change Buyout Process to Ready To Invoice
	if (String(Request.Form("Processed"))=="True") {
		var rsSetBuyoutStatus = Server.CreateObject("ADODB.Recordset");
		rsSetBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetBuyoutStatus.Source = "{call dbo.cp_update_buyout_status("+Request.QueryString("intBuyout_req_id")+",4,0)}";
		rsSetBuyoutStatus.CursorType = 0;
		rsSetBuyoutStatus.CursorLocation = 2;
		rsSetBuyoutStatus.LockType = 3;
		rsSetBuyoutStatus.Open();	

		var rsSetBuyoutProcess = Server.CreateObject("ADODB.Recordset");
		rsSetBuyoutProcess.ActiveConnection = MM_cnnASP02_STRING;
		rsSetBuyoutProcess.Source = "update tbl_buyout_request set insBuyout_Prc_id = 4 where intBuyout_req_id = " + Request.QueryString("intBuyout_req_id");
		rsSetBuyoutProcess.CursorType = 0;
		rsSetBuyoutProcess.CursorLocation = 2;
		rsSetBuyoutProcess.LockType = 3;
		rsSetBuyoutProcess.Open();	
	}
	
	//Trigger to change Buyout Status to Buyout Delivered	
	//Trigger to insert into client/institution services and notes page with I-Tech code
	if (String(Request.Form("Delivered"))=="True") {
		var rsSetBuyoutStatus = Server.CreateObject("ADODB.Recordset");
		rsSetBuyoutStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetBuyoutStatus.Source = "{call dbo.cp_update_buyout_status("+Request.QueryString("intBuyout_req_id")+",3,0)}";
		rsSetBuyoutStatus.CursorType = 0;
		rsSetBuyoutStatus.CursorLocation = 2;
		rsSetBuyoutStatus.LockType = 3;
		rsSetBuyoutStatus.Open();				
		
		Notes = "Loan Request: " + Request.QueryString("intBuyout_req_id") + " shipped.\n" + Notes;	
		var Year = CurrentYear();
		var Cycle = CurrentMonth();
		//User Type
		switch (String(rsBuyout.Fields.Item("insEq_user_type").Value)) {
			//Client
			case "3":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsBuyout.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','3200000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();
			break;
			//Institution
			case "4":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+rsBuyout.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','2F00000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();			
			break;
		}
	}		
	Response.Redirect("m010e0501.asp?intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));	
}

var rsMethod = Server.CreateObject("ADODB.Recordset");
rsMethod.ActiveConnection = MM_cnnASP02_STRING;

// + Nov.04.2005
//rsMethod.Source = "{call dbo.cp_buyout_ship_method3("+ intShip_dtl_id + ",0,0,'',0,0,0,'',0,0,'','',0,0,'Q',0)}";
rsMethod.Source = "{call dbo.cp_buyout_ship_method("+ intShip_dtl_id + ",0,'',0,0,0,'',0,'','',0,0,'Q',0)}";
                                                                    
rsMethod.CursorType = 0;
rsMethod.CursorLocation = 2;
rsMethod.LockType = 3;
rsMethod.Open();

var IsNew = ((rsMethod.EOF)?true:false);

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsShippingMethod = Server.CreateObject("ADODB.Recordset");
rsShippingMethod.ActiveConnection = MM_cnnASP02_STRING;
rsShippingMethod.Source = "{call dbo.cp_shipping_method(0,0)}";
rsShippingMethod.CursorType = 0;
rsShippingMethod.CursorLocation = 2;
rsShippingMethod.LockType = 3;
rsShippingMethod.Open();

rsNotes.Close();
rsNotes.Open();

var rsBoxes = Server.CreateObject("ADODB.Recordset");
rsBoxes.ActiveConnection = MM_cnnASP02_STRING;
rsBoxes.Source = "{call dbo.cp_buyout_ship_box(0,"+intShip_dtl_id+",0,0,0,0,0,'Q',0)}";
rsBoxes.CursorType = 0;
rsBoxes.CursorLocation = 2;
rsBoxes.LockType = 3;
rsBoxes.Open();
var total = 0
while (!rsBoxes.EOF) {
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
				document.frm0501.reset();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Init(){
		switch (document.frm0501.ShippingMethod.value) {
			//dynamex
			case "9":
				document.frm0501.PickedUp[0].disabled = false;
				document.frm0501.PickedUp[1].disabled = false;	
				document.frm0501.WayBillNumber.disabled = false;
			break;
			//picked up by client
			case "10":
				document.frm0501.PickedUp[0].disabled = false;
				document.frm0501.PickedUp[1].disabled = false;
				document.frm0501.WayBillNumber.disabled = true;
			break;
			//taken by consultant
			case "1":
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = true;												
			break;
			//loomis
			case "4":
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = false;
			break;
			//none
			default:			
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = true;
			break;
		}
		document.frm0501.IsProcessed.focus();
	}

	function openWindow(page){
		if (page!='nothing') win1=window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function ListBoxes(){	
		openWindow('m010pop2.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&intShip_dtl_id=<%=intShip_dtl_id%>');		
	}
	
	function ChangeShippingMethod(){
		switch (document.frm0501.ShippingMethod.value) {
			//dynamex
			case "9":
				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";
				document.frm0501.PickedUp[0].disabled = false;
				document.frm0501.PickedUp[1].disabled = false;	
				document.frm0501.WayBillNumber.disabled = false;
			break;
			//picked up by client
			case "10":
				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0501.PickedUp[0].disabled = false;
				document.frm0501.PickedUp[1].disabled = false;
				document.frm0501.WayBillNumber.disabled = true;
			break;
			//taken by consultant
			case "1":
				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = true;												
			break;
			//loomis
			case "4":
				document.frm0501.ScheduledArrivalDate.value=ForwardDay(1);
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = false;
			break;
			//none
			default:			
				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = true;
			break;
		}
	}
	
	function Save(){
		if (!CheckTextArea(document.frm0501.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}	
//		if ((!CheckDate(document.frm0501.DateProcessed.value)) || (Trim(document.frm0501.DateProcessed.value)=="")){
		if (!CheckDate(document.frm0501.DateProcessed.value)){
			alert("Invalid Date Processed.");
			return ;
		}
		if (!CheckDate(document.frm0501.DeliveryDate.value)){
			alert("Invalid Delivery Date.");
			return ;
		}
		if (!CheckDate(document.frm0501.ScheduledArrivalDate.value)){
			alert("Invalid Scheduled Arrival Date.");
			document.frm0501.ScheduledArrivalDate.focus();
			return ;
		}
		
		if (!document.frm0501.IsProcessed.checked) document.frm0501.Processed.value="False";
		if (!document.frm0501.IsDelivered.checked) document.frm0501.Delivered.value="False";
						
		if (Trim(document.frm0501.NumberOfBoxes.value)=="") document.frm0501.NumberOfBoxes.value="0";
		document.frm0501.submit();
	}
	
	function ChangeProcessed() {
		if (document.frm0501.IsProcessed.checked) {
			document.frm0501.DateProcessed.disabled = false;			
			document.frm0501.DateProcessed.value = "<%=CurrentDate()%>";
		} else {
			document.frm0501.DateProcessed.disabled = true;
			document.frm0501.DateProcessed.value = "";
		}		
	}	

	function ChangeDelivered() {
		if (document.frm0501.IsDelivered.checked) {
			document.frm0501.DeliveryDate.disabled = false;			
			document.frm0501.DeliveryDate.value = "<%=CurrentDate()%>";
		} else {
			document.frm0501.DeliveryDate.disabled = true;
			document.frm0501.DeliveryDate.value = "";
		}		
	}	
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0501">
<h5>Shipping Method</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap><input type="checkbox" name="IsProcessed" tabindex="1" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("bitIs_Process").Value==1)?"CHECKED":""))%> onClick="ChangeProcessed();" accesskey="F" class="chkstyle">Date Processed:</td>
		<td nowrap>
			<input type="text" name="DateProcessed" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsUser_Ship_date").Value):"")%>" tabindex="2" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>  
	<tr> 
		<td nowrap>Processed By:</td>
		<td nowrap><select name="ProcessedBy" tabindex="3">
			<option value="0">(none)		
		<% 
		while (!rsStaff.EOF) {
		%>
			<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%if (!IsNew) { Response.Write(((rsStaff.Fields.Item("insStaff_id").Value==rsMethod.Fields.Item("insShip_Staff_id").Value)?"SELECTED":"")) } else { Response.Write(((rsStaff.Fields.Item("insStaff_id").Value==Session("insStaff_id"))?"SELECTED":""))}%>><%=(rsStaff.Fields.Item("chvName").Value)%></option>
		<%
			rsStaff.MoveNext();
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Shipping Method:</td>
		<td nowrap><select name="ShippingMethod" onChange="ChangeShippingMethod();" tabindex="4">
			<option value="0">(none)
	<% 
	while (!rsShippingMethod.EOF) {
		if (rsShippingMethod.Fields.Item("bitis_active").Value == "1") {
	%>
			<option value="<%=(rsShippingMethod.Fields.Item("intship_method_id").Value)%>" <%if (!IsNew) Response.Write(((rsShippingMethod.Fields.Item("intship_method_id").Value==rsMethod.Fields.Item("insShip_Method_id").Value)?"SELECTED":""))%>><%=(rsShippingMethod.Fields.Item("chvname").Value)%></option>
	<%
		}
		rsShippingMethod.MoveNext();
	}
	%>
		</select></td>
    </tr>	
	<tr> 
		<td nowrap>Waybill Number:</td>
		<td nowrap><input type="text" name="WayBillNumber" size="15" value="<%=((!IsNew)?(rsMethod.Fields.Item("chvWayBill_No").Value):"")%>" tabindex="5"></td>
	</tr>
	<tr>
		<td nowrap>Number of Boxes:</td>
		<td nowrap>
			<input type="text" name="NumberOfBoxes" size="2" maxlength="3" value="<%=((!IsNew)?rsMethod.Fields.Item("insNum_of_Boxes").Value:0)%>" tabindex="6" style="border: none" readOnly onKeypres="AllowNumericOnly();"> Total Weight: 
			<input type="text" name="TotalWeight" size="4" value="<%=((!IsNew)?total:"0")%>" tabindex="7" style="border: none" readOnly> LB 
			<input type="button" value="Add/Update" tabindex="8" onClick="<%=((!IsNew)?"ListBoxes();":"alert('Please save first, before adding shipping boxes.');")%>" class="btnstyle">
		</td>
	</tr>
    <tr>
		<td nowrap><input type="checkbox" name="IsDelivered" tabindex="9" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("bitIs_Dlvy").Value==1)?"CHECKED":""))%> onClick="ChangeDelivered();" class="chkstyle">Delivery Date:</td>
		<td nowrap>
			<input type="text" name="DeliveryDate" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsDlvy_date").Value):"")%>" tabindex="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
			<!--<input type="text" name="DeliveryDayOfTheWeek" size="20" value="" readonly style="border: none">-->
		</td>
    </tr>
    <tr>
		<td nowrap>Scheduled Arrival Date:</td>
		<td nowrap>
			<input type="text" name="ScheduledArrivalDate" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsSch_Arv_date").Value):"")%>" tabindex="11" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Picked Up:</td>
		<td nowrap>
			<input type="radio" name="PickedUp" value="1" tabindex="12" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value=="1")?"CHECKED":""))%> class="chkstyle">Morning
			<input type="radio" name="PickedUp" value="0" tabindex="13" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value=="0")?"CHECKED":""))%> class="chkstyle">Afternoon
		</td>
	</tr>
	<tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap><textarea name="Notes" cols="65" rows="3" tabindex="14" accesskey="L"><%=((!rsNotes.EOF)?rsNotes.Fields.Item("chvNote_Desc").Value:"")%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="15" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="16" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=rsBuyout.Fields.Item("intShip_dtl_id").Value %>">
<input type="hidden" name="Processed" value="<%if (!IsNew) {Response.Write((rsMethod.Fields.Item("bitIs_Process").Value=="1")?"False":"True")} else {Response.Write("True")}%>">
<input type="hidden" name="Delivered" value="<%if (!IsNew) {Response.Write((rsMethod.Fields.Item("bitIs_Dlvy").Value=="1")?"False":"True")} else {Response.Write("True")}%>">
</form>
</body>
</html>
<%
rsMethod.Close();
rsStaff.Close();
rsShippingMethod.Close();
rsNotes.Close();
%>