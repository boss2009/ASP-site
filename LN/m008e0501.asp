<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_editAction += "?" + Request.QueryString;
}

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var intShip_dtl_id = 0;
if (!rsLoan.EOF) {
	if (rsLoan.Fields.Item("intShip_dtl_id").Value != null) intShip_dtl_id = rsLoan.Fields.Item("intShip_dtl_id").Value;
} 

var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_loan_ship_notes("+ intShip_dtl_id + ",0,'',"+Session("insStaff_id")+",0,'Q',0)}";
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
	var DateProcessed = ((String(Request.Form("DateProcessed"))!="undefined")?Request.Form("DateProcessed"):"1/1/1900");	
	var IsProcessed = ((Request.Form("IsProcessed")=="on")?"1":"0");		
	var rsMethod = Server.CreateObject("ADODB.Recordset");
	rsMethod.ActiveConnection = MM_cnnASP02_STRING;
	rsMethod.Source = "{call dbo.cp_loan_ship_method2("+Request.Form("MM_recordId")+","+Request.QueryString("intLoan_Req_id")+","+IsProcessed+",'"+DateProcessed+"',"+Request.Form("ProcessedBy")+","+Request.Form("ShippingMethod")+","+intShip_notes_id+",'"+WayBillNumber+"',"+Request.Form("NumberOfBoxes")+",'"+Request.Form("DeliveryDate")+"','"+Request.Form("ScheduledArrivalDate")+"',"+MorningPickedUp+",0,'E',0)}";
	rsMethod.CursorType = 0;
	rsMethod.CursorLocation = 2;
	rsMethod.LockType = 3;
	rsMethod.Open();

	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	rsNotes.Close();
	if (intShip_notes_id != 0) {		
		rsNotes.Source = "{call dbo.cp_loan_ship_notes(0,"+intShip_notes_id+",'"+Notes+"',"+Session("insStaff_id")+",0,'E',0)}";	
	} else {
		rsNotes.Source = "{call dbo.cp_loan_ship_notes("+intShip_dtl_id+","+intShip_notes_id+",'"+Notes+"',"+Session("insStaff_id")+",0,'A',0)}";	
	}
	rsNotes.Open();		

	//Trigger for changing Loan Status to Loan Processed and all Inventory Status to on-Loan, User ID and Type to Loan User ID and Type.
	if (String(Request.Form("Processed"))=="True") {
		var rsSetLoanStatus = Server.CreateObject("ADODB.Recordset");
		rsSetLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetLoanStatus.Source = "{call dbo.cp_update_loan_status("+Request.QueryString("intLoan_req_id")+",3,0)}";
		rsSetLoanStatus.CursorType = 0;
		rsSetLoanStatus.CursorLocation = 2;
		rsSetLoanStatus.LockType = 3;
		rsSetLoanStatus.Open();			

		var SetToStatus = 3;
		
		switch (String(rsLoan.Fields.Item("insLoan_Type_id").Value)) {
			case "6":
				SetToStatus = 21;
			break;		
			case "1":
				SetToStatus = 2;
			break;		
			case "3":
				SetToStatus = 4;
			break;		
			case "15":
				SetToStatus = 4;
			break;		
			case "10":
				SetToStatus = 4;
			break;		
			case "12":
				SetToStatus = 4;
			break;		
			case "11":
				SetToStatus = 4;
			break;		
			case "4":
				SetToStatus = 20;
			break;		
			case "5":
				SetToStatus = 19;
			break;		
			case "2":
				SetToStatus = 3;
			break;		
			case "7":
				SetToStatus = 26;
			break;		
			default:
				SetToStatus = 3;
			break;		
		}
		
		var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_Req_id")+",0,'',0,0,'','',0,'Q',0)}";
		rsInventoryLoaned.CursorType = 0;
		rsInventoryLoaned.CursorLocation = 2;
		rsInventoryLoaned.LockType = 3;
		rsInventoryLoaned.Open();
		while (!rsInventoryLoaned.EOF) {
			//only update if the status of the inventory is allocated
			if (String(rsInventoryLoaned.Fields.Item("insCurrent_Status").Value)=="25") {			
				var rsSetInventoryStatus = Server.CreateObject("ADODB.Recordset");
				rsSetInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
				rsSetInventoryStatus.Source = "{call dbo.cp_Update_eqpIvtry_status("+rsInventoryLoaned.Fields.Item("intEquip_set_id").Value+","+SetToStatus+",0)}";
				rsSetInventoryStatus.CursorType = 0;
				rsSetInventoryStatus.CursorLocation = 2;
				rsSetInventoryStatus.LockType = 3;
				rsSetInventoryStatus.Open();		
				
				if (String(rsLoan.Fields.Item("insEq_user_type").Value)=="4") {			
					var rsSetInventoryUser = Server.CreateObject("ADODB.Recordset");
					rsSetInventoryUser.ActiveConnection = MM_cnnASP02_STRING;
					rsSetInventoryUser.Source = "update tbl_equip_inventory set insInstit_User_id = " + rsLoan.Fields.Item("intEq_user_id").Value + ", insUser_Type_id = 4 where intEquip_set_id = " + rsInventoryLoaned.Fields.Item("intEquip_set_id").Value;
					rsSetInventoryUser.CursorType = 0;
					rsSetInventoryUser.CursorLocation = 2;
					rsSetInventoryUser.LockType = 3;
					rsSetInventoryUser.Open();		
				} else {
					var rsSetInventoryUser = Server.CreateObject("ADODB.Recordset");
					rsSetInventoryUser.ActiveConnection = MM_cnnASP02_STRING;
					rsSetInventoryUser.Source = "update tbl_equip_inventory set insUser_id = " + rsLoan.Fields.Item("intEq_user_id").Value + ", insUser_Type_id = " + rsLoan.Fields.Item("insEq_user_type").Value + " where intEquip_set_id = " + rsInventoryLoaned.Fields.Item("intEquip_set_id").Value;
					rsSetInventoryUser.CursorType = 0;
					rsSetInventoryUser.CursorLocation = 2;
					rsSetInventoryUser.LockType = 3;
					rsSetInventoryUser.Open();		
				}
			}			
			rsInventoryLoaned.MoveNext();			 
		}	
	}
		
	//Trigger to insert into client/institution services and notes page with I-Tech code
	//Trigger to change loan status to Loan Delivered		
	if (String(Request.Form("Delivered"))=="True") {
		Notes = "Loan Request: " + Request.QueryString("intLoan_req_id") + " shipped.\n" + Notes;	
		var Year = CurrentYear();
		var Cycle = CurrentMonth();
		//User Type
		switch (String(rsLoan.Fields.Item("insEq_user_type").Value)) {
			//Client
			case "3":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsLoan.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','2E00000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();
			break;
			//Institution
			case "4":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+rsLoan.Fields.Item("insInst_User_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','2F00000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();			
			break;
		}
		
		var rsSetLoanStatus = Server.CreateObject("ADODB.Recordset");
		rsSetLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetLoanStatus.Source = "{call dbo.cp_update_loan_status("+Request.QueryString("intLoan_req_id")+",5,0)}";
		rsSetLoanStatus.CursorType = 0;
		rsSetLoanStatus.CursorLocation = 2;
		rsSetLoanStatus.LockType = 3;
		rsSetLoanStatus.Open();					
	}
		
	Response.Redirect("m008e0501.asp?intLoan_Req_id="+Request.QueryString("intLoan_Req_id"));
}

if (String(Request("MM_action")) == "insert") {
	var WayBillNumber = ((String(Request.Form("WayBillNumber"))!="undefined")?String(Request.Form("WayBillNumber")).replace(/'/g, "''"):"");			
	var MorningPickedUp = null;
	if (String(Request.Form("PickedUp"))!="undefined") MorningPickedUp = Request.Form("PickedUp");
	var DateProcessed = ((String(Request.Form("DateProcessed"))!="undefined")?Request.Form("DateProcessed"):"1/1/1900");		
	var IsProcessed = ((Request.Form("IsProcessed")=="on")?"1":"0");		
	var cmdInsertShipDetail = Server.CreateObject("ADODB.Command");
	cmdInsertShipDetail.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertShipDetail.CommandText = "dbo.cp_loan_Ship_Method2";
	cmdInsertShipDetail.CommandType = 4;
	cmdInsertShipDetail.CommandTimeout = 0;
	cmdInsertShipDetail.Prepared = true;
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intRecID", 3, 1,1,0));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intLoan_req_id", 3, 1,1,Request.QueryString("intLoan_req_id")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@bitIs_Process", 2, 1,1,IsProcessed));		
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@dtsUser_Ship_date", 200, 1,30,DateProcessed));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insShip_Staff_id", 2, 1,1,Request.Form("ProcessedBy")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insShip_Method_id", 2, 1,1,Request.Form("ShippingMethod")));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@intShip_notes_id", 3, 1,1,intShip_notes_id));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@chvWayBill_No", 200, 1,20,WayBillNumber));
	cmdInsertShipDetail.Parameters.Append(cmdInsertShipDetail.CreateParameter("@insNum_of_Boxes", 2, 1,1,Request.Form("NumberOfBoxes")));
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
	rsNotes.Source = "{call dbo.cp_loan_ship_notes("+ cmdInsertShipDetail.Parameters.Item("@intRtnValue").Value + ",0,'"+Notes+"',"+Session("insStaff_id")+",0,'A',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();	
	
	//Trigger for changing Loan Status to Loan Processed and all Inventory Status to on-Loan, User ID and Type to Loan User ID and Type.
	if (String(Request.Form("Processed"))=="True") {
		var rsSetLoanStatus = Server.CreateObject("ADODB.Recordset");
		rsSetLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetLoanStatus.Source = "{call dbo.cp_update_loan_status("+Request.QueryString("intLoan_req_id")+",3,0)}";
		rsSetLoanStatus.CursorType = 0;
		rsSetLoanStatus.CursorLocation = 2;
		rsSetLoanStatus.LockType = 3;
		rsSetLoanStatus.Open();			

		var SetToStatus = 3;
		
		switch (String(rsLoan.Fields.Item("insLoan_Type_id").Value)) {
			case "6":
				SetToStatus = 21;
			break;		
			case "1":
				SetToStatus = 2;
			break;		
			case "3":
				SetToStatus = 4;
			break;		
			case "15":
				SetToStatus = 4;
			break;		
			case "10":
				SetToStatus = 4;
			break;		
			case "12":
				SetToStatus = 4;
			break;		
			case "11":
				SetToStatus = 4;
			break;		
			case "4":
				SetToStatus = 20;
			break;		
			case "5":
				SetToStatus = 19;
			break;		
			case "2":
				SetToStatus = 3;
			break;		
			case "7":
				SetToStatus = 26;
			break;		
			default:
				SetToStatus = 3;
			break;		
		}
		
		var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.QueryString("intLoan_req_id")+",0,'',0,0,'','',0,'Q',0)}";
		rsInventoryLoaned.CursorType = 0;
		rsInventoryLoaned.CursorLocation = 2;
		rsInventoryLoaned.LockType = 3;
		rsInventoryLoaned.Open();
		while (!rsInventoryLoaned.EOF) {
			//only update if the status of the inventory is allocated
			if (String(rsInventoryLoaned.Fields.Item("insCurrent_Status").Value)=="25") {		
				var rsSetInventoryStatus = Server.CreateObject("ADODB.Recordset");
				rsSetInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
				rsSetInventoryStatus.Source = "{call dbo.cp_Update_eqpIvtry_status("+rsInventoryLoaned.Fields.Item("intEquip_set_id").Value+","+SetToStatus+",0)}";
				rsSetInventoryStatus.CursorType = 0;
				rsSetInventoryStatus.CursorLocation = 2;
				rsSetInventoryStatus.LockType = 3;
				rsSetInventoryStatus.Open();
	
				if (String(rsLoan.Fields.Item("insEq_user_type").Value)=="4") {			
					var rsSetInventoryUser = Server.CreateObject("ADODB.Recordset");
					rsSetInventoryUser.ActiveConnection = MM_cnnASP02_STRING;
					rsSetInventoryUser.Source = "update tbl_equip_inventory set insInstit_User_id = " + rsLoan.Fields.Item("intEq_user_id").Value + ", insUser_Type_id = 4 where intEquip_set_id = " + rsInventoryLoaned.Fields.Item("intEquip_set_id").Value;
					rsSetInventoryUser.CursorType = 0;
					rsSetInventoryUser.CursorLocation = 2;
					rsSetInventoryUser.LockType = 3;
					rsSetInventoryUser.Open();		
				} else {
					var rsSetInventoryUser = Server.CreateObject("ADODB.Recordset");
					rsSetInventoryUser.ActiveConnection = MM_cnnASP02_STRING;
					rsSetInventoryUser.Source = "update tbl_equip_inventory set insUser_id = " + rsLoan.Fields.Item("intEq_user_id").Value + ", insUser_Type_id = " + rsLoan.Fields.Item("insEq_user_type").Value + " where intEquip_set_id = " + rsInventoryLoaned.Fields.Item("intEquip_set_id").Value;
					rsSetInventoryUser.CursorType = 0;
					rsSetInventoryUser.CursorLocation = 2;
					rsSetInventoryUser.LockType = 3;
					rsSetInventoryUser.Open();		
				}
			}					
			rsInventoryLoaned.MoveNext();
		}	
	}
		
	//Trigger to insert into client/institution services and notes page with I-Tech code
	//Trigger to change loan status to Loan Delivered		
	if (String(Request.Form("Delivered"))=="True") {
		Notes = "Loan Request: " + Request.QueryString("intLoan_req_id") + " shipped.\n" + Notes;	
		var Year = CurrentYear();
		var Cycle = CurrentMonth();
		//User Type
		switch (String(rsLoan.Fields.Item("insEq_user_type").Value)) {
			//Client
			case "3":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_ac_srv_note("+rsLoan.Fields.Item("intEq_user_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','2E00000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();
			break;
			//Institution
			case "4":
				var rsServiceRequested = Server.CreateObject("ADODB.Recordset");
				rsServiceRequested.ActiveConnection = MM_cnnASP02_STRING;
				rsServiceRequested.Source = "{call dbo.cp_pilat_srv_note("+rsLoan.Fields.Item("insInst_User_id").Value+",0,'"+Request.Form("DateProcessed")+"',"+Year+","+Cycle+","+Session("insStaff_id")+",'"+Notes+"','2F00000000000000000000000000000000000000',0,'A',0)}";
				rsServiceRequested.CursorType = 0;
				rsServiceRequested.CursorLocation = 2;
				rsServiceRequested.LockType = 3;
				rsServiceRequested.Open();			
			break;
		}
		
		var rsSetLoanStatus = Server.CreateObject("ADODB.Recordset");
		rsSetLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
		rsSetLoanStatus.Source = "{call dbo.cp_update_loan_status("+Request.QueryString("intLoan_req_id")+",5,0)}";
		rsSetLoanStatus.CursorType = 0;
		rsSetLoanStatus.CursorLocation = 2;
		rsSetLoanStatus.LockType = 3;
		rsSetLoanStatus.Open();					
	}
		
	Response.Redirect("m008e0501.asp?intLoan_Req_id="+Request.QueryString("intLoan_Req_id"));	
}

var rsMethod = Server.CreateObject("ADODB.Recordset");
rsMethod.ActiveConnection = MM_cnnASP02_STRING;

//+ Nov.03.2005
//rsMethod.Source = "{call dbo.cp_loan_ship_method2("+ intShip_dtl_id + ",0,0,'',0,0,0,'',0,'','',0,0,'Q',0)}";
rsMethod.Source = "{call dbo.cp_loan_ship_method("+ intShip_dtl_id + ",0,'',0,0,0,'',0,'','',0,0,'Q',0)}";

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
rsBoxes.Source = "{call dbo.cp_loan_ship_box(0,"+intShip_dtl_id+",0,0,0,0,0,'Q',0)}";
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
		document.frm0501.IsProcessed.focus();	
		ChangeShippingMethod();
	}

	function openWindow(page){
		if (page!='nothing') win1=window.open(page, "", "width=300,height=300,scrollbars=1,left=300,top=300,status=1");
		return ;
	}
	
	function ChangeShippingMethod(){
		switch (document.frm0501.ShippingMethod.value) {
			//dynamex
			case "9":
//				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";
				document.frm0501.PickedUp[0].disabled = false;
				document.frm0501.PickedUp[1].disabled = false;	
				document.frm0501.WayBillNumber.disabled = false;
			break;
			//picked up by client
			case "10":
//				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0501.PickedUp[0].disabled = false;
				document.frm0501.PickedUp[1].disabled = false;
				document.frm0501.WayBillNumber.disabled = true;
			break;
			//taken by consultant
			case "1":
//				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = true;												
			break;
			//loomis
			case "4":
//				document.frm0501.ScheduledArrivalDate.value=ForwardDay(1);
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = false;
			break;
			//none
			default:			
//				document.frm0501.ScheduledArrivalDate.value="<%=CurrentDate()%>";			
				document.frm0501.PickedUp[0].disabled = true;
				document.frm0501.PickedUp[1].disabled = true;
				document.frm0501.WayBillNumber.disabled = true;
			break;
		}
	}
	
	function ListBoxes(){	
		openWindow('m008pop2.asp?intLoan_Req_id=<%=Request.QueryString("intLoan_Req_id")%>&intShip_dtl_id=<%=intShip_dtl_id%>');		
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
		
		if (!document.frm0501.IsProcessed.checked) document.frm0501.Processed.value="False";
		
		if (!CheckDate(document.frm0501.DeliveryDate.value)){
			alert("Invalid Delivery Date.");
			document.frm0501.DeliveryDate.focus();
			return ;
		}
		
		if ((Trim(document.frm0501.DeliveryDate.value)!="") && (document.frm0501.Delivered.value=="True")) {
			document.frm0501.Delivered.value = "True";
		} else {
			document.frm0501.Delivered.value = "False";
		}
		
		if (!CheckDate(document.frm0501.ScheduledArrivalDate.value)){
			alert("Invalid Scheduled Arrival Date.");
			document.frm0501.ScheduledArrivalDate.focus();
			return ;
		}
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
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0501">
<h5>Shipping Method</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
<!-- + Nov.03.2005
-->
		<td nowrap><input type="checkbox" name="IsProcessed" tabindex="1" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value==1)?"CHECKED":""))%> onClick="ChangeProcessed();" accesskey="F" class="chkstyle">Date Processed:</td>

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
			<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>" <%if (!IsNew) { Response.Write(((rsStaff.Fields.Item("insStaff_id").Value==rsMethod.Fields.Item("insShip_Staff_id").Value)?"SELECTED":""))} else { Response.Write(((rsStaff.Fields.Item("insStaff_id").Value==Session("insStaff_id"))?"SELECTED":""))}%>><%=(rsStaff.Fields.Item("chvName").Value)%></option>
		<%
			rsStaff.MoveNext();
		}
		%>
        </select></td>
	</tr>
	<tr> 
		<td nowrap>Shipping Method:</td>
		<td nowrap><select name="ShippingMethod" tabindex="4" onChange="ChangeShippingMethod();">
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
			<input type="text" name="NumberOfBoxes" size="2" maxlength="3" value="<%=((!IsNew)?rsMethod.Fields.Item("insNum_of_Boxes").Value:0)%>" style="border: none" tabindex="6" readonly onKeypres="AllowNumericOnly();">
			Total Weight: <input type="text" name="TotalWeight" size="4" value="<%=((!IsNew)?total:"0")%>" style="border: none" readonly tabindex="7">
			LB <input type="button" value="Add/Update" onClick="<%=((!IsNew)?"ListBoxes();":"alert('Please save first, before adding shipping boxes.');")%>" tabindex="8" class="btnstyle">
		</td>		
	</tr>
	<tr> 
		<td nowrap>Delivery Date:</td>
		<td nowrap>
			<input type="text" name="DeliveryDate" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsDlvy_date").Value):"")%>" tabindex="9" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
    </tr>
    <tr> 
		<td nowrap>Scheduled Arrival Date:</td>
		<td nowrap>
			<input type="text" name="ScheduledArrivalDate" size="11" maxlength="10" value="<%=((!IsNew)?FilterDate(rsMethod.Fields.Item("dtsSch_Arv_date").Value):"")%>" tabindex="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>		
    </tr>
	<tr>
		<td nowrap>Picked Up:</td>
		<td nowrap>
			<input type="radio" name="PickedUp" value="1" tabindex="11" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value=="1")?"CHECKED":""))%> class="chkstyle">Morning
			<input type="radio" name="PickedUp" value="0" tabindex="12" <%if (!IsNew) Response.Write(((rsMethod.Fields.Item("BitPkup_morning").Value=="0")?"CHECKED":""))%> class="chkstyle">Afternoon
		</td>
	</tr>
	<tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="3" tabindex="13" accesskey="L"><%=((!rsNotes.EOF)?rsNotes.Fields.Item("chvNote_Desc").Value:"")%></textarea></td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="14" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="15" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_action" value="<%=((IsNew)?"insert":"update")%>">
<input type="hidden" name="MM_recordId" value="<%=rsLoan.Fields.Item("intShip_dtl_id").Value %>">
<!-- + Nov.03.2005
-->
<input type="hidden" name="Processed" value="<%if (!IsNew) {Response.Write((rsMethod.Fields.Item("BitPkup_morning").Value=="1")?"False":"True")} else {Response.Write("True")}%>">

<input type="hidden" name="Delivered" value="<%if (!IsNew) {Response.Write(((String(rsMethod.Fields.Item("dtsDlvy_date").Value)=="Mon Jan 1 00:00:00 PST 1900")||(rsMethod.Fields.Item("dtsDlvy_date").Value==null))?"True":"False")} else {Response.Write("True")}%>">
</form>
</body>
</html>
<%
rsMethod.Close();
rsStaff.Close();
rsShippingMethod.Close();
rsNotes.Close();
%>