<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var EquipUserID
	var IsBackOrdered = ((Request.Form("EquipmentOnBackOrder")=="on") ? "1":"0");
	var InstUserID
	var IsIdvUser

	switch (String(Request.Form("UserType"))) {
		//none
		case "0":
			EquipUserID = 0;
			InstUserID = 0;
			IsIdvUser = 0;
		break;
		//staff
		case "1":		
			EquipUserID = Request.Form("IndividualUserID");
			InstUserID = 0;
			IsIdvUser = 1;
		break;
		//client
		case "3":
			EquipUserID = Request.Form("IndividualUserID");
			InstUserID = 0;
			IsIdvUser = 1;
		break;
		//institution
		case "4":
			EquipUserID = Request.Form("InstitutionUserID");
			InstUserID = Request.Form("InstitutionUserID");
			IsIdvUser = 0;
		break;
		default:
			EquipUserID = 0;
			InstUserID = 0;
			IsIdvUser = 0;
		break;
	}
	var Year = ((Request.Form("Year")!="")?Request.Form("Year"):"0");
	var Cycle = ((Request.Form("Cycle")!="")?Request.Form("Cycle"):"0");	
	var rsLoan = Server.CreateObject("ADODB.Recordset");
	rsLoan.ActiveConnection = MM_cnnASP02_STRING;
	rsLoan.Source = "{call dbo.cp_loan_request2("+Request.Form("MM_recordId")+","+Request.Form("LoanType")+","+EquipUserID+","+Request.Form("UserType")+","+Session("insStaff_id")+",'"+Request.Form("DateRequested")+"',"+Request.Form("ApprovedBy")+",'"+Request.Form("DateApproved")+"',"+Request.Form("LoanStatus")+","+Year+","+Cycle+","+IsBackOrdered+","+InstUserID+","+Request.Form("DurationOfLoan")+","+Request.Form("DurationPeriod")+",'"+Request.Form("LoanDueDate")+"',"+IsIdvUser+","+Session("insStaff_id")+",0,0,'',0,'E',0)}";
	rsLoan.CursorType = 0;
	rsLoan.CursorLocation = 2;
	rsLoan.LockType = 3;
	rsLoan.Open();
	
	if (String(Request.Form("CancelLoan"))=="True") {
		var rsInventoryLoaned = Server.CreateObject("ADODB.Recordset");
		rsInventoryLoaned.ActiveConnection = MM_cnnASP02_STRING;
		rsInventoryLoaned.Source = "{call dbo.cp_eqp_loaned(0,"+Request.Form("MM_recordId")+",0,'',0,0,'','',0,'Q',0)}";
		rsInventoryLoaned.CursorType = 0;
		rsInventoryLoaned.CursorLocation = 2;
		rsInventoryLoaned.LockType = 3;
		rsInventoryLoaned.Open();
		while (!rsInventoryLoaned.EOF) {
			var rsSetInventoryStatus = Server.CreateObject("ADODB.Recordset");
			rsSetInventoryStatus.ActiveConnection = MM_cnnASP02_STRING;
			rsSetInventoryStatus.Source = "{call dbo.cp_Update_eqpIvtry_status("+rsInventoryLoaned.Fields.Item("intEquip_set_id").Value+",1,0)}";
			rsSetInventoryStatus.CursorType = 0;
			rsSetInventoryStatus.CursorLocation = 2;
			rsSetInventoryStatus.LockType = 3;
			rsSetInventoryStatus.Open();		
			rsInventoryLoaned.MoveNext();
		}
	}
	
	Response.Redirect("m008e0101.asp?intLoan_Req_id="+Request.Form("MM_recordId"));
}

var rsLoanType = Server.CreateObject("ADODB.Recordset");
rsLoanType.ActiveConnection = MM_cnnASP02_STRING;
rsLoanType.Source = "{call dbo.cp_loan_type2(0,'',0,0,'Q',0)}";
rsLoanType.CursorType = 0;
rsLoanType.CursorLocation = 2;
rsLoanType.LockType = 3;
rsLoanType.Open();

var rsLoanStatus = Server.CreateObject("ADODB.Recordset");
rsLoanStatus.ActiveConnection = MM_cnnASP02_STRING;
rsLoanStatus.Source = "{call dbo.cp_loan_status(0,0)}";
rsLoanStatus.CursorType = 0;
rsLoanStatus.CursorLocation = 2;
rsLoanStatus.LockType = 3;
rsLoanStatus.Open();

var rsDurationType = Server.CreateObject("ADODB.Recordset");
rsDurationType.ActiveConnection = MM_cnnASP02_STRING;
rsDurationType.Source = "{call dbo.cp_duratn_type2(0,'','',0,'Q',0)}";
rsDurationType.CursorType = 0;
rsDurationType.CursorLocation = 2;
rsDurationType.LockType = 3;
rsDurationType.Open();

var rsUserType = Server.CreateObject("ADODB.Recordset");
rsUserType.ActiveConnection = MM_cnnASP02_STRING;
rsUserType.Source = "{call dbo.cp_eq_user_type2(0,'',1,0,0,'Q',0)}";
rsUserType.CursorType = 0;
rsUserType.CursorLocation = 2;
rsUserType.LockType = 3;
rsUserType.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();

var rsLoan = Server.CreateObject("ADODB.Recordset");
rsLoan.ActiveConnection = MM_cnnASP02_STRING;
rsLoan.Source = "{call dbo.cp_loan_request2("+ Request.QueryString("intLoan_Req_id") + ",0,0,0,0,'',0,'',0,0,0,0,0,0,0,'',0,0,1,0,'',1,'Q',0)}";
rsLoan.CursorType = 0;
rsLoan.CursorLocation = 2;
rsLoan.LockType = 3;
rsLoan.Open();

var InstUserName = "";
var InstUserId = 0;
var IdvUserName = "";
var IdvUserId = 0;
switch (String(rsLoan.Fields.Item("insEq_user_type").Value)){
	//staff
	case "1":		
		var rsIndStaff = Server.CreateObject("ADODB.Recordset");
		rsIndStaff.ActiveConnection = MM_cnnASP02_STRING;
		rsIndStaff.Source = "{call dbo.cp_staff("+rsLoan.Fields.Item("intEq_user_id").Value+",0,1)}";
		rsIndStaff.CursorType = 0;
		rsIndStaff.CursorLocation = 2;
		rsIndStaff.LockType = 3;
		rsIndStaff.Open();
		if (!rsIndStaff.EOF) {
			IdvUserName = rsIndStaff.Fields.Item("chvLst_Name").Value + ", " +rsIndStaff.Fields.Item("chvFst_Name").Value ;		
		} else {
			IdvUserName = "";
		}
		IdvUserId = rsLoan.Fields.Item("intEq_user_id").Value;
		rsIndStaff.Close();
	break;
	//client
	case "3":		
		var rsIndClient = Server.CreateObject("ADODB.Recordset");
		rsIndClient.ActiveConnection = MM_cnnASP02_STRING;
		rsIndClient.Source = "{call dbo.cp_Idv_Adult_Client("+rsLoan.Fields.Item("intEq_user_id").Value+")}";
		rsIndClient.CursorType = 0;
		rsIndClient.CursorLocation = 2;
		rsIndClient.LockType = 3;
		rsIndClient.Open();
		if (!rsIndClient.EOF) {
			IdvUserName = rsIndClient.Fields.Item("chvLst_Name").Value + ", " + rsIndClient.Fields.Item("chvFst_Name").Value;		
		} else {
			IdvUserName = "";
		}
		IdvUserId = rsLoan.Fields.Item("intEq_user_id").Value;		
		rsIndClient.Close();
	break;
	//institution
	case "4":
		var rsIndInstitution = Server.CreateObject("ADODB.Recordset");
		rsIndInstitution.ActiveConnection = MM_cnnASP02_STRING;		
		rsIndInstitution.Source = "{call dbo.cp_school3("+rsLoan.Fields.Item("intEq_user_id").Value+",'',0,0,0,0,0,0,0,'',1,'Q',0)}";
		rsIndInstitution.CursorType = 0;
		rsIndInstitution.CursorLocation = 2;
		rsIndInstitution.LockType = 3;
		rsIndInstitution.Open();
		InstUserName = ((!rsIndInstitution.EOF)?rsIndInstitution.Fields.Item("chvSchool_Name").Value:"");
		InstUserId = rsLoan.Fields.Item("intEq_user_id").Value;		
		rsIndInstitution.Close();		
	break;
	default:
	break;
}
%>									
<html>
<head>
	<title>General Information</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83:
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0101.reset();
			break;
		}
	}
	</script>	
	<script language="Javascript">	
	function ChangeLoanType(){
		switch (document.frm0101.LoanType.value){
			//none
			case "0":
				RemoveOption();
				AddOption('0','(none)');
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDuration.style.visibility = "hidden";
				oLoanDueDateLabel.style.visibility = "hidden";
				oLoanDueDate.style.visibility = "hidden";

				oIndividualUserLabel.style.visibility = "hidden";
				document.frm0101.IndividualUserName.style.visibility = "hidden";				
				document.frm0101.ListIndividualUser.style.visibility = "hidden";
				
				oInstitutionUserLabel.style.visibility = "hidden";				
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
				
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";
			break;
			//short term
			case "1":
				RemoveOption();
				AddOption('3','Client');			
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";																				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";																

				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";				
			break;
			//long term
			case "2":
				RemoveOption();
				AddOption('3','Client');						
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDuration.style.visibility = "hidden";
				oLoanDueDateLabel.style.visibility = "hidden";
				oLoanDueDate.style.visibility = "hidden";

				oIndividualUserLabel.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";												
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
				
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//assessment
			case "3":
				RemoveOption();
				AddOption('3','Client');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "visible";				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";								
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
																
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//interim
			case "4":
				RemoveOption();
				AddOption('4','Institution');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "hidden";				
				document.frm0101.ListIndividualUser.style.visibility = "hidden";
				document.frm0101.IndividualUserName.style.visibility = "hidden";								
				
				oInstitutionUserLabel.style.visibility = "visible";								
				document.frm0101.ListInstitutionUser.style.visibility = "visible";
				document.frm0101.InstitutionUserName.style.visibility = "visible";
																				
				oPILATReferralDateLabel.style.visibility = "visible";
				oPILATReferralDateReminder.style.visibility = "visible";																
				document.frm0101.PILATReferralDate.style.visibility = "visible";				
			break;
			//low utilization		
			case "5":
				RemoveOption();
				AddOption('4','Institution');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";
				
				oIndividualUserLabel.style.visibility = "hidden";				
				document.frm0101.ListIndividualUser.style.visibility = "hidden";
				document.frm0101.IndividualUserName.style.visibility = "hidden";								
				
				oInstitutionUserLabel.style.visibility = "visible";								
				document.frm0101.ListInstitutionUser.style.visibility = "visible";
				document.frm0101.InstitutionUserName.style.visibility = "visible";
																
				oPILATReferralDateLabel.style.visibility = "visible";
				oPILATReferralDateReminder.style.visibility = "visible";																
				document.frm0101.PILATReferralDate.style.visibility = "visible";								
			break;
			//staff
			case "6":
				RemoveOption();
				AddOption('1','Staff');						
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDuration.style.visibility = "hidden";
				oLoanDueDateLabel.style.visibility = "hidden";
				oLoanDueDate.style.visibility = "hidden";

				oIndividualUserLabel.style.visibility = "visible";				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";				
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
																
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//employment
			case "7":
				RemoveOption();
				AddOption('3','Client');						
				oLoanDuration.style.visibility = "hidden";
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDueDate.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				
				oIndividualUserLabel.style.visibility = "visible";				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";				
								
				oInstitutionUserLabel.style.visibility = "hidden";												
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";

				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																												
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;	
			//standard
			default:
				RemoveOption();
				AddOption('3','Client');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";												
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
				
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			
		}
		
		document.frm0101.InstitutionUserName.value = "";
		document.frm0101.InstitutionUserID.value = 0;
		document.frm0101.IndividualUserName.value = "";
		document.frm0101.IndividualUserID.value = 0;		
	}
	
	function ChangeLoanDuration(){
		if (Trim(document.frm0101.DurationOfLoan.value)!=""){
			switch (document.frm0101.DurationPeriod.value) {
				case "0":
					document.frm0101.LoanDueDate.value="";
				break;
				case "3":
					document.frm0101.LoanDueDate.value=ForwardDay(document.frm0101.DurationOfLoan.value);
				break;
				case "4":
					document.frm0101.LoanDueDate.value=ForwardWeek(document.frm0101.DurationOfLoan.value);
				break;
				case "5":
					document.frm0101.LoanDueDate.value=ForwardMonth(document.frm0101.DurationOfLoan.value);
				break;
				case "6":
					document.frm0101.LoanDueDate.value=ForwardYear(document.frm0101.DurationOfLoan.value);
				break;
			}
		}
	}
	
	function RemoveOption(){
	  	while (document.frm0101.UserType.length > 0){
			document.frm0101.UserType.remove(0);
  		}	
	}
	
	function AddOption(val, txt){
		var oOption=document.createElement("OPTION");
		oOption.text = txt;
		oOption.value = val;
		document.frm0101.UserType.add(oOption);
	}
	
	function Init(){
		switch (document.frm0101.LoanType.value){
			//none
			case "0":
				RemoveOption();
				AddOption('0','(none)');
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDuration.style.visibility = "hidden";
				oLoanDueDateLabel.style.visibility = "hidden";
				oLoanDueDate.style.visibility = "hidden";

				oIndividualUserLabel.style.visibility = "hidden";
				document.frm0101.IndividualUserName.style.visibility = "hidden";				
				document.frm0101.ListIndividualUser.style.visibility = "hidden";
				
				oInstitutionUserLabel.style.visibility = "hidden";				
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
				
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";
			break;
			//short term
			case "1":
				RemoveOption();
				AddOption('3','Client');			
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";																				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";																

				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";				
			break;
			//long term
			case "2":
				RemoveOption();
				AddOption('3','Client');						
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDuration.style.visibility = "hidden";
				oLoanDueDateLabel.style.visibility = "hidden";
				oLoanDueDate.style.visibility = "hidden";

				oIndividualUserLabel.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";												
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
				
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//assessment
			case "3":
				RemoveOption();
				AddOption('3','Client');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "visible";				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";								
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
																
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//interim
			case "4":
				RemoveOption();
				AddOption('4','Institution');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "hidden";				
				document.frm0101.ListIndividualUser.style.visibility = "hidden";
				document.frm0101.IndividualUserName.style.visibility = "hidden";								
				
				oInstitutionUserLabel.style.visibility = "visible";								
				document.frm0101.ListInstitutionUser.style.visibility = "visible";
				document.frm0101.InstitutionUserName.style.visibility = "visible";
																				
				oPILATReferralDateLabel.style.visibility = "visible";
				oPILATReferralDateReminder.style.visibility = "visible";																
				document.frm0101.PILATReferralDate.style.visibility = "visible";				
			break;
			//low utilization		
			case "5":
				RemoveOption();
				AddOption('4','Institution');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";
				
				oIndividualUserLabel.style.visibility = "hidden";				
				document.frm0101.ListIndividualUser.style.visibility = "hidden";
				document.frm0101.IndividualUserName.style.visibility = "hidden";								
				
				oInstitutionUserLabel.style.visibility = "visible";								
				document.frm0101.ListInstitutionUser.style.visibility = "visible";
				document.frm0101.InstitutionUserName.style.visibility = "visible";
																
				oPILATReferralDateLabel.style.visibility = "visible";
				oPILATReferralDateReminder.style.visibility = "visible";																
				document.frm0101.PILATReferralDate.style.visibility = "visible";								
			break;
			//staff
			case "6":
				RemoveOption();
				AddOption('1','Staff');						
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDuration.style.visibility = "hidden";
				oLoanDueDateLabel.style.visibility = "hidden";
				oLoanDueDate.style.visibility = "hidden";

				oIndividualUserLabel.style.visibility = "visible";				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";				
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
																
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//employment
			case "7":
				RemoveOption();
				AddOption('3','Client');						
				oLoanDuration.style.visibility = "hidden";
				oLoanDurationLabel.style.visibility = "hidden";			
				oLoanDueDate.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				
				oIndividualUserLabel.style.visibility = "visible";				
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";				
								
				oInstitutionUserLabel.style.visibility = "hidden";												
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";

				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																												
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;
			//Default
			default:
				RemoveOption();
				AddOption('3','Client');						
				oLoanDurationLabel.style.visibility = "visible";			
				oLoanDuration.style.visibility = "visible";
				oLoanDueDateLabel.style.visibility = "visible";
				oLoanDueDate.style.visibility = "visible";

				oIndividualUserLabel.style.visibility = "visible";
				document.frm0101.IndividualUserName.style.visibility = "visible";												
				document.frm0101.ListIndividualUser.style.visibility = "visible";
				
				oInstitutionUserLabel.style.visibility = "hidden";								
				document.frm0101.ListInstitutionUser.style.visibility = "hidden";
				document.frm0101.InstitutionUserName.style.visibility = "hidden";
				
				oPILATReferralDateLabel.style.visibility = "hidden";
				oPILATReferralDateReminder.style.visibility = "hidden";																																
				document.frm0101.PILATReferralDate.style.visibility = "hidden";								
			break;				
		}
		document.frm0101.DateRequested.focus();
	}

	function ChangeStatus(){
		if (document.frm0101.LoanStatus.value=="2") {
			if (Trim(document.frm0101.DateApproved.value)=="") {
				document.frm0101.DateApproved.value = "<%=CurrentDate()%>";
			}
			if (document.frm0101.ApprovedBy.value <= 0) {
				document.frm0101.ApprovedBy.value = "<%=Session("insStaff_id")%>";
			}
		}
	}
	
	function Save(){
		if (!CheckDate(document.frm0101.DateRequested.value)){
			alert("Invalid Date Requested.");
			document.frm0101.DateRequested.focus();
			return ;
		}
		if (!(document.frm0101.LoanType.value > "0")){
			alert("Select a Loan Type.");
			document.frm0101.LoanType.focus();
			return ;
		}		
		if ((document.frm0101.CancelLoan.value=="True") && (document.frm0101.LoanStatus.value=="6")) {
			document.frm0101.CancelLoan.value = "True";
		} else {
			document.frm0101.CancelLoan.value = "False";
		}
		document.frm0101.submit();
	}

	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=460,height=430,scrollbars=0,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_editAction%>" method="POST" name="frm0101">
<h5>General Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Date Requested:</td>
		<td nowrap width="180">
			<input type="text" name="DateRequested" value="<%=FilterDate(rsLoan.Fields.Item("dtsRequest_date").Value)%>" size="11" maxlength="10" readonly tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Year/Cycle:</td>
		<td nowrap>
			<input type="text" name="Year" value="<%=GetYear(rsLoan.Fields.Item("intYear_Cycle").Value)%>" size="4" maxlength="4" tabindex="2" onKeypress="AllowNumericOnly();">&nbsp;
			<input type="text" name="Cycle" value="<%=GetCycle(rsLoan.Fields.Item("intYear_Cycle").Value)%>" size="2" maxlength="2" tabindex="3" onKeypress="AllowNumericOnly();">
		</td>
	</tr>
	<tr>
		<td nowrap>Loan Status:</td>
		<td nowrap colspan="3"><select name="LoanStatus" onChange="ChangeStatus();" tabindex="4">
			<%
			while (!rsLoanStatus.EOF) {
			%>
				<option value="<%=(rsLoanStatus.Fields.Item("intloan_status_id").Value)%>" <%=((rsLoanStatus.Fields.Item("intloan_status_id").Value==rsLoan.Fields.Item("insLoan_Status_id").Value)?"SELECTED":"")%>><%=(rsLoanStatus.Fields.Item("chvname").Value)%>
			<%
				rsLoanStatus.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td nowrap>Date Approved:</td>
		<td nowrap>
			<input type="text" name="DateApproved" value="<%=FilterDate(rsLoan.Fields.Item("dtsApprvd_Date").Value)%>" readonly size="11" maxlength="10" tabindex="5" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Approved By:</td>
		<td nowrap><select name="ApprovedBy" tabindex="6">
				<option value="0">(none)
			<%
			while (!rsStaff.EOF) {
			%>
				<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsLoan.Fields.Item("insApprvd_Staff_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%>
			<%
				rsStaff.MoveNext();
			}
			rsStaff.MoveFirst();
			%>	
		</select></td>
	</tr>
	<tr>
		<td nowrap>Loan Type:</td>
		<td nowrap><select name="LoanType" tabindex="7" onChange="ChangeLoanType();">
			<%
			while (!rsLoanType.EOF) {
			%>
				<option value="<%=(rsLoanType.Fields.Item("intloan_type_id").Value)%>" <%=((rsLoanType.Fields.Item("intloan_type_id").Value==rsLoan.Fields.Item("insLoan_Type_id").Value)?"SELECTED":"")%>><%=(rsLoanType.Fields.Item("chvname").Value)%>
			<%
				rsLoanType.MoveNext();
			}
			%>
		</select></td>
		<td nowrap><div id="oLoanDurationLabel">Duration of Loan:</div></td>
		<td nowrap><div id="oLoanDuration">
			<input type="text" name="DurationOfLoan" value="<%=(rsLoan.Fields.Item("insLoan_Duration").Value)%>" size="3" onChange="ChangeLoanDuration();" onKeypress="AllowNumericOnly();" maxlength="3" tabindex="13">
			<select name="DurationPeriod" onChange="ChangeLoanDuration();" tabindex="14">
			<%
			while (!rsDurationType.EOF) {
			%>
				<option value="<%=(rsDurationType.Fields.Item("insDuratn_type_id").Value)%>" <%=((rsDurationType.Fields.Item("insDuratn_type_id").Value==rsLoan.Fields.Item("insDuratn_type_id").Value)?"SELECTED":"")%>><%=(rsDurationType.Fields.Item("chrAbbrev").Value)%>
			<%
				rsDurationType.MoveNext();
			}
			%>
			</select>
		</div></td>				
	</tr>
	<tr> 
		<td nowrap>User Type:</td>
		<td nowrap><select name="UserType" tabindex="8">
			<%
			while (!rsUserType.EOF) {
			%>
				<option value="<%=(rsUserType.Fields.Item("insEq_user_type").Value)%>" <%=((rsUserType.Fields.Item("insEq_user_type").Value==rsLoan.Fields.Item("insEq_user_type").Value)?"SELECTED":"")%>><%=(rsUserType.Fields.Item("chvEq_user_type").Value)%></option>
			<%
				rsUserType.MoveNext();
			}
			%>
		</select></td>				
		<td nowrap><div id="oLoanDueDateLabel">Loan Due Date:</div></td>
		<td nowrap><div id="oLoanDueDate">
			<input type="text" name="LoanDueDate" value="<%=FilterDate(rsLoan.Fields.Item("dtsLoan_Due_Date").Value)%>" size="11" maxlength="10" tabindex="15" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</div></td>
	</tr>
	<tr>			
		<td nowrap><div id="oIndividualUserLabel">Individual User:</div></td>
		<td nowrap colspan="3">
			<input type="text" name="IndividualUserName" value="<%=IdvUserName%>" readonly tabindex="9">
			<input type="button" name="ListIndividualUser" value="List" onClick="if (document.frm0101.UserType.value == '1') {openWindow('m008p0201.asp','wPopUser');} else {openWindow('m008p0202.asp','wPopUser');}" tabindex="10" class="btnstyle">
		</td>
    </tr>
    <tr> 
		<td nowrap><div id="oInstitutionUserLabel">Institution User:</div></td>
		<td nowrap>
			<input type="text" name="InstitutionUserName" value="<%=InstUserName%>" readonly tabindex="11">
			<input type="button" name="ListInstitutionUser" value="List" onClick="openWindow('m008p0301.asp','wPopUser');" tabindex="12" class="btnstyle">
		</td>
		<td nowrap><div id="oPILATReferralDateLabel">PILAT Referral Date:</div></td>
		<td nowrap>
			<input type="text" name="PILATReferralDate" size="11" maxlength="10" tabindex="16" onChange="FormatDate(this)">
			<span id="oPILATReferralDateReminder" style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap colspan="4"><input type="checkbox" name="EquipmentOnBackOrder" <%=((rsLoan.Fields.Item("bitIsBack_Ordered").Value=="1")?"CHECKED":"")%> tabindex="17" readonly accesskey="L" class="chkstyle">Equipment on Backorder</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="18" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="19" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="InstitutionUserID" value="<%=InstUserId%>">
<input type="hidden" name="IndividualUserID" value="<%=IdvUserId%>">
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="CancelLoan" value="<%=((rsLoan.Fields.Item("insLoan_Status_id").Value==6)?"False":"True")%>">
<input type="hidden" name="MM_recordId" value="<%=rsLoan.Fields.Item("intLoan_Req_id").Value %>">
</form>
</body>
</html>
<%
rsLoanType.Close();
rsLoanStatus.Close();
rsDurationType.Close();
rsUserType.Close();
rsLoan.Close();
%>