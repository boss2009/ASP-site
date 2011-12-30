<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
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
			IsIdvUser =1 ;
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
	var DurationOfLoan = ((Request.Form("DurationOfLoan")!="")?Request.Form("DurationOfLoan"):"0");
	var DurationPeriod = ((Request.Form("DurationPeriod")!="")?Request.Form("DurationPeriod"):"0");	
	var LoanDueDate = ((Request.Form("LoanDueDate")!="")?Request.Form("LoanDueDate"):"1/1/1900");
	var DateApproved = "1/1/1900";
	var DateRequested = ((Request.Form("DateRequested")!="")?Request.Form("DateRequested"):"1/1/1900");		
			
	var cmdInsertLoan = Server.CreateObject("ADODB.Command");
	cmdInsertLoan.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertLoan.CommandText = "dbo.cp_Loan_Request2";
	cmdInsertLoan.CommandType = 4;
	cmdInsertLoan.CommandTimeout = 0;
	cmdInsertLoan.Prepared = true;
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@intRecId", 3, 1,1,0));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insLoan_Type_id", 2, 1,1,Request.Form("LoanType")));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@intEq_user_id", 3, 1,1,EquipUserID));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insEq_user_type", 2, 1,1,Request.Form("UserType")));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insReqst_Staff_id", 2, 1,1,Session("insStaff_id")));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@dtsRequest_date", 135, 1,1,DateRequested));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insApprvd_Staff_id", 2, 1,1,0));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@dtsApprvd_Date", 135, 1,1,DateApproved));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insLoan_Status_id", 2, 1,1,Request.Form("LoanStatus")));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insYear", 2, 1,1,Year));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insCycle", 16, 1,1,Cycle));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@bitIsBack_Ordered", 2, 1,1,IsBackOrdered));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insInst_User_id", 2, 1,1,InstUserID));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insLoan_Duration", 2, 1,1,DurationOfLoan));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insDuratn_type_id", 2, 1,1,DurationPeriod));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@dtsLoan_Due_Date", 135, 1,1,LoanDueDate));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@bitIs_Idv_User", 2, 1,1,IsIdvUser));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insUser_id", 2, 1,1,Session("insStaff_id")));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@inspSrtBy", 2, 1,1,0));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@inspSrtOrd", 2, 1,1,0));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@chvFilter", 200, 1,1,''));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@insMode", 16, 1,1,0));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@chvTask", 129, 1,1,'A'));
	cmdInsertLoan.Parameters.Append(cmdInsertLoan.CreateParameter("@intRtnFlag", 3, 2));
	cmdInsertLoan.Execute();
	
	Response.Redirect("m008FS3.asp?intLoan_req_id="+cmdInsertLoan.Parameters.Item("@intRtnFlag").Value);
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
%>									
<html>
<head>
	<title>New Loan Request</title>
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
		   	case 76 :
				//alert("L");
				self.close();
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
			//Long Term
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
			//default
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
	
	function Save(){
		if ((document.frm0101.InstitutionUserID.value == 0) && (document.frm0101.IndividualUserID.value == 0)) {
			alert("Select a user.");
			return ;
		}
		
		if (!(document.frm0101.LoanType.value > "0")){
			alert("Select a Loan Type.");
			document.frm0101.LoanType.focus();
			return ;
		}		
		if (!CheckDate(document.frm0101.DateRequested.value)){
			alert("Invalid Date Requested.");
			document.frm0101.DateRequested.focus();
			return ;
		}
		document.frm0101.submit();
		document.frm0101.btnSave.disabled = true;		
	}

	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=460,height=430,scrollbars=0,left=0,top=0,status=1");
		return ;
	}	
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_editAction%>" method="POST" name="frm0101">
<h5>New Loan Request</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Date Requested:</td>
		<td nowrap>
			<input type="text" name="DateRequested" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
		<td nowrap>Year/Cycle:</td>
		<td nowrap>
			<input type="text" name="Year" value="<%=CurrentYear()%>" size="4" maxlength="4" tabindex="2" onKeypress="AllowNumericOnly();">&nbsp;
			<input type="text" name="Cycle" value="<%=ZeroPadFormat(CurrentMonth(),2)%>" size="2" maxlength="2" tabindex="2" onKeypress="AllowNumericOnly();">
		</td>
	</tr>
	<tr>
		<td nowrap>Loan Status:</td>
		<td nowrap><select name="LoanStatus" tabindex="5">
			<%
			while (!rsLoanStatus.EOF) {
			%>
				<option value="<%=(rsLoanStatus.Fields.Item("intloan_status_id").Value)%>" <%=((rsLoanStatus.Fields.Item("intloan_status_id").Value=="1")?"SELECTED":"")%>><%=(rsLoanStatus.Fields.Item("chvname").Value)%>
			<%
				rsLoanStatus.MoveNext();
			}
			%>
		</select></td>
		<td colspan="2"></td>
	</tr>
	<tr>
		<td nowrap>Loan Type:</td>
		<td nowrap><select name="LoanType" tabindex="6" onChange="ChangeLoanType();">
			<%
			while (!rsLoanType.EOF) {
			%>
				<option value="<%=(rsLoanType.Fields.Item("intloan_type_id").Value)%>" <%=((rsLoanType.Fields.Item("intloan_type_id").Value=="8")?"SELECTED":"")%>><%=(rsLoanType.Fields.Item("chvname").Value)%>
			<%
				rsLoanType.MoveNext();
			}
			%>
		</select></td>
		<td nowrap><div id="oLoanDurationLabel">Duration of Loan:</div></td>
		<td nowrap><div id="oLoanDuration">
			<input type="text" name="DurationOfLoan" size="3" onChange="ChangeLoanDuration();" onKeypress="AllowNumericOnly();" maxlength="3" tabindex="12">
			<select name="DurationPeriod" onChange="ChangeLoanDuration();" tabindex="13">
			<%
			while (!rsDurationType.EOF) {
			%>
				<option value="<%=(rsDurationType.Fields.Item("insDuratn_type_id").Value)%>"><%=(rsDurationType.Fields.Item("chrAbbrev").Value)%>
			<%
				rsDurationType.MoveNext();
			}
			%>
			</select>
		</div></td>				
	</tr>
	<tr> 
		<td nowrap>User Type:</td>
		<td nowrap><select name="UserType" tabindex="7">
			<%
			while (!rsUserType.EOF) {
			%>
				<option value="<%=(rsUserType.Fields.Item("insEq_user_type").Value)%>"><%=(rsUserType.Fields.Item("chvEq_user_type").Value)%></option>
			<%
				rsUserType.MoveNext();
			}
			%>
		</select></td>				
		<td nowrap><div id="oLoanDueDateLabel">Loan Due Date:</div></td>
		<td nowrap><div id="oLoanDueDate">
			<input type="text" name="LoanDueDate" size="11" maxlength="10" tabindex="14" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</div></td>
	</tr>
	<tr>			
		<td nowrap><div id="oIndividualUserLabel">Individual User:</div></td>
		<td nowrap>
			<input type="text" name="IndividualUserName" readonly tabindex="8">
			<input type="button" name="ListIndividualUser" value="List" onClick="if (document.frm0101.UserType.value == '1') {openWindow('m008p0201.asp','wPopUser');} else {openWindow('m008p0202.asp','wPopUser');}" tabindex="9" class="btnstyle">
		</td>
		<td colspan="2"></td>
    </tr>
    <tr> 
		<td nowrap><div id="oInstitutionUserLabel">Institution User:</div></td>
		<td nowrap>
			<input type="text" name="InstitutionUserName" readonly tabindex="10">
			<input type="button" name="ListInstitutionUser" value="List" onClick="openWindow('m008p0301.asp','wPopUser');" tabindex="11" class="btnstyle">
		</td>
		<td nowrap><div id="oPILATReferralDateLabel">PILAT Referral Date:</div></td>
		<td nowrap>
			<input type="text" name="PILATReferralDate" size="11" maxlength="10" tabindex="15" onChange="FormatDate(this)">
			<span id="oPILATReferralDateReminder" style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td colspan="4"><input type="checkbox" name="EquipmentOnBackOrder" tabindex="16" accesskey="L" class="chkstyle">Equipment on Backorder</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" name="btnSave" value="Save" tabindex="17" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Cancel" tabindex="18" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="IndividualUserID">
<input type="hidden" name="InstitutionUserID">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsLoanType.Close();
rsLoanStatus.Close();
rsDurationType.Close();
rsUserType.Close();
%>