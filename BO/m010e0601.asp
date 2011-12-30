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

var intBOShip_dtl_id = 0;
if (!rsBuyout.EOF) {
	if (rsBuyout.Fields.Item("intBOShip_dtl_id").Value != null) intBOShip_dtl_id = rsBuyout.Fields.Item("intBOShip_dtl_id").Value;
} 

var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_buyout_ship_notes("+ intBOShip_dtl_id + ",0,'',"+Session("insStaff_id")+",0,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();

var intShip_notes_id = 0;
if (!rsNotes.EOF) {
	if (rsNotes.Fields.Item("intShip_notes_id").Value != null) intShip_notes_id = rsNotes.Fields.Item("intShip_notes_id").Value;
}

if (String(Request("MM_action")) == "update") {
	var rsMethod = Server.CreateObject("ADODB.Recordset");
	rsMethod.ActiveConnection = MM_cnnASP02_STRING;
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))!="undefined")?String(Request.Form("WayBillNumber")).replace(/'/g, "''"):"");			
	var MorningPickedUp = null;
	if (String(Request.Form("PickedUp"))!="undefined") MorningPickedUp = Request.Form("PickedUp");
	rsMethod.Source = "{call dbo.cp_buyout_ship_method("+Request.Form("MM_recordId")+","+Request.QueryString("intBuyout_req_id")+",'"+Request.Form("DateProcessed")+"',"+Request.Form("ProcessedBy")+","+Request.Form("ShippingMethod")+","+intShip_notes_id+",'"+WayBillNumber+"',"+Request.Form("NumberOfBoxes")+",'"+Request.Form("DeliveryDate")+"','"+Request.Form("ScheduledArrivalDate")+"',"+MorningPickedUp+",1,'E',0)}";
	rsMethod.CursorType = 0;
	rsMethod.CursorLocation = 2;
	rsMethod.LockType = 3;
	rsMethod.Open();

	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	rsNotes.Close();
	if (intShip_notes_id != 0) {
		rsNotes.Source = "{call dbo.cp_buyout_ship_notes(0,"+intShip_notes_id+",'"+Notes+"',"+Session("insStaff_id")+",0,'E',0)}";	
	} else {
		rsNotes.Source = "{call dbo.cp_buyout_ship_notes("+intBOShip_dtl_id+","+intShip_notes_id+",'"+Notes+"',"+Session("insStaff_id")+",0,'A',0)}";	
	}
	rsNotes.Open();		
	Response.Redirect("m010e0601.asp?intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));
}

if (String(Request("MM_action")) == "insert") {
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))!="undefined")?String(Request.Form("WayBillNumber")).replace(/'/g, "''"):"");			
	var MorningPickedUp = null;
	if (String(Request.Form("PickedUp"))!="undefined") MorningPickedUp = Request.Form("PickedUp");
	var cmdInsertShipDetail = Server.CreateObject("ADODB.Command");
	cmdInsertShipDetail.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertShipDetail.CommandText = "dbo.cp_Buyout_Ship_Method";
	cmdInsertShipDetail.CommandType = 4;
	cmdInsertShipDetail.CommandTimeout = 0;
	cmdInsertShipDetail.Prepared = true;
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intRecID", 3, 1,1,0));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intBuyout_Req_id", 3, 1,1,Request.QueryString("intBuyout_req_id")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsUser_Ship_date", 200, 1,30,Request.Form("DateProcessed")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insShip_Staff_id", 2, 1,1,Request.Form("ProcessedBy")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insShip_Method_id", 2, 1,1,Request.Form("ShippingMethod")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intShip_notes_id", 3, 1,1,intShip_notes_id));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@chvWayBill_No", 200, 1,20,WayBillNumber));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insNum_of_Boxes", 2, 1,1,Request.Form("NumberOfBoxes")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsDlvy_date", 200, 1,30,Request.Form("DeliveryDate")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsSch_Arv_date", 200, 1,30,Request.Form("ScheduledArrivalDate")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@BitPkup_morning", 2, 1,1,MorningPickedUp));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insMode", 16, 1,1,1));
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
	Response.Redirect("m010e0601.asp?intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));	
}

var rsMethod = Server.CreateObject("ADODB.Recordset");
rsMethod.ActiveConnection = MM_cnnASP02_STRING;
rsMethod.Source = "{call dbo.cp_buyout_ship_method("+ intBOShip_dtl_id + ",0,'',0,0,0,'',0,'','',0,0,'Q',0)}";
rsMethod.CursorType = 0;
rsMethod.CursorLocation = 2;
rsMethod.LockType = 3;
rsMethod.Open();

var IsNew = ((!rsMethod.EOF)?false:true);

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
rsBoxes.Source = "{call dbo.cp_buyout_ship_box(0,"+intBOShip_dtl_id+",0,0,0,0,0,'Q',0)}";
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
	<title>Backorder Shipping Method</title>
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
				document.frm0601.reset();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Init(){
		ChangeShippingMethod();
		document.frm0601.DateProcessed.focus();
	}

	function openWindow(page){
		if (page!='nothing') win1=window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function ListBoxes(){	
		openWindow('m010pop3.asp?intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>&intBOShip_dtl_id=<%=intBOShip_dtl_id%>');
	}
	
	function ChangeShippingMethod(){
		switch (document.frm0601.ShippingMethod.value) {
			//dynamex
			case "9":
//				document.frm0601.ScheduledArrivalDate.value="<%=CurrentDate()%>";
				document.frm0601.PickedUp[0].disabled = false;
				document.frm0601.PickedUp[1].disabled = false;	
				document.frm0601.WayBillNumber.disabled = false;
			break;
			//picked up by client
			case "10":
//				document.frm0601.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0601.PickedUp[0].disabled = false;
				document.frm0601.PickedUp[1].disabled = false;
				document.frm0601.WayBillNumber.disabled = true;
			break;
			//taken by consultant
			case "1":
//				document.frm0601.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0601.PickedUp[0].disabled = true;
				document.frm0601.PickedUp[1].disabled = true;
				document.frm0601.WayBillNumber.disabled = true;												
			break;
			//loomis
			case "4":
//				document.frm0601.ScheduledArrivalDate.value=ForwardDay(1);
				document.frm0601.PickedUp[0].disabled = true;
				document.frm0601.PickedUp[1].disabled = true;
				document.frm0601.WayBillNumber.disabled = false;
			break;
			//none
			default:			
//				document.frm0601.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0601.PickedUp[0].disabled = true;
				document.frm0601.PickedUp[1].disabled = true;
				document.frm0601.WayBillNumber.disabled = true;
			break;
		}
	}
	
	function Save(){
		if (!CheckTextArea(document.frm0601.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}	
		if (!CheckDate(document.frm0601.DateProcessed.value)){
			alert("Invalid Date Processed.");
			document.frm0601.DateProcessed.focus();
			return ;
		}
		if (!CheckDate(document.frm0601.DeliveryDate.value)){
			alert("Invalid Delivery Date.");
			document.frm0601.DeliveryDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0601.ScheduledArrivalDate.value)){
			alert("Invalid Scheduled Arrival Date.");
			document.frm0601.ScheduledArrivalDate.focus();
			return ;
		}
		if (Trim(document.frm0601.NumberOfBoxes.value)=="") document.frm0601.NumberOfBoxes.value="0";
						
		document.frm0601.submit();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0601">
<h5>Backorder Shipping Method</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Date Processed:</td>
		<td nowrap>
			<input type="text" name="DateProcessed" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsUser_Ship_date").Value):CurrentDate())%>" tabindex="1" accesskey="F" onChange="FormatDate(this)">
        	<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
    <tr> 
		<td nowrap>Processed By:</td>
		<td nowrap><select name="ProcessedBy" tabindex="2">
			<option value="0">(none)		
		<% 
		while (!rsStaff.EOF) {
		%>
			<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%if (!IsNew) { Response.Write(((rsMethod.Fields.Item("insShip_Staff_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":""))} else { Response.Write(((rsStaff.Fields.Item("insStaff_id").Value==Session("insStaff_id"))?"SELECTED":""))};%>><%=(rsStaff.Fields.Item("chvName").Value)%></option>
		<%
			rsStaff.MoveNext();
		}
		%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Shipping Method:</td>
		<td nowrap><select name="ShippingMethod" tabindex="3" onChange="ChangeShippingMethod();">
			<option value="0">(none)		
	<% 
	while (!rsShippingMethod.EOF) {
		if (rsShippingMethod.Fields.Item("bitis_active").Value == "1") {
	%>
			<option value="<%=(rsShippingMethod.Fields.Item("intship_method_id").Value)%>" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("insShip_Method_id").Value==rsShippingMethod.Fields.Item("intship_method_id").Value)?"SELECTED":""));%>><%=(rsShippingMethod.Fields.Item("chvname").Value)%></option>
	<%
		}
		rsShippingMethod.MoveNext();
	}
	%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Waybill Number:</td>
		<td nowrap><input type="text" name="WayBillNumber" size="15" value="<%=((!IsNew)?rsMethod.Fields.Item("chvWayBill_No").Value:"")%>" tabindex="4"></td>
    </tr>
    <tr> 
		<td nowrap>Number of Boxes:</td>
		<td nowrap>
			<input type="text" name="NumberOfBoxes" size="2" maxlength="3" value="<%=((!IsNew)?rsMethod.Fields.Item("insNum_of_Boxes").Value:0)%>" tabindex="5" style="border: none" readonly onKeypres="AllowNumericOnly();">
			Total Weight: <input type="text" name="TotalWeight" size="4" value="<%=((!IsNew)?total:"0")%>" tabindex="6" style="border: none" readonly>
			LB <input type="button" value="Add/Update" tabindex="7" onClick="<%=((!IsNew)?"ListBoxes();":"alert('Please save first, before adding shipping boxes.');")%>" class="btnstyle">
		</td>
    </tr>
    <tr> 
		<td nowrap>Delivery Date:</td>
		<td nowrap>
			<input type="text" name="DeliveryDate" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsDlvy_date").Value):"")%>" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
    <tr> 
		<td nowrap>Scheduled Arrival Date:</td>
		<td nowrap>
			<input type="text" name="ScheduledArrivalDate" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsSch_Arv_date").Value):"")%>" tabindex="9" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Picked Up:</td>
		<td nowrap>
			<input type="radio" name="PickedUp" value="1" tabindex="10" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value=="1")?"CHECKED":""))%> class="chkstyle">Morning 
        	<input type="radio" name="PickedUp" value="0" tabindex="11" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value=="0")?"CHECKED":""))%> class="chkstyle">Afternoon 
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="12" accesskey="L"><%=((!rsNotes.EOF)?rsNotes.Fields.Item("chvNote_Desc").Value:"")%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="13" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="14" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=intBOShip_dtl_id%>">
</form>
</body>
</html>
<%
rsMethod.Close();
rsStaff.Close();
rsShippingMethod.Close();
rsNotes.Close();
%>