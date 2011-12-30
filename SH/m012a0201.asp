<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_insertAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_insertAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true"){
	var LowUtilization = 0;
	var Interim = 0;
	var TechConsult = 0;
	var ConsultationOnly = 0;
	var LoanReferral = 0;
	var GrantReferral = 0;
	var OtherReferral = 0;				
	var HardwareSoftwarePurchasedByInstitution = 0;
	var PurchaseOrderSigned = 0;
	var LoanRequiredBy = "";
	var LoanReturned = "";	

	switch (String(Request.Form("ReferralType"))) {
		case "1":
			LowUtilization = 1;
			LoanReferral = 1;
		break;
		case "2":
			Interim = 1;
			LoanReferral = 1;
			HardwareSoftwarePurchasedByInstitution = ((Request.Form("HardwareSoftwarePurchasedByInstitution")=="on")?"1":"0");				
			PurchaseOrderSigned = ((Request.Form("PurchaseOrderSigned")=="on")?"1":"0");							
			LoanRequiredBy = Request.Form("LoanRequiredBy");
			LoanReturned = Request.Form("LoanReturned");
		break;
		case "3":
			TechConsult = 1;
			GrantReferral = 1;
		break;
		case "4":
			ConsultationOnly = 1;
			OtherReferral = 1;
		break;
	}

	var rsPILATReferral = Server.CreateObject("ADODB.Recordset");
	rsPILATReferral.ActiveConnection = MM_cnnASP02_STRING;
	rsPILATReferral.Source = "{call dbo.cp_PILAT_referrals3("+Request.QueryString("insSchool_id")+",0,'"+Request.Form("ReferralDate")+"',"+HardwareSoftwarePurchasedByInstitution+","+PurchaseOrderSigned+",'"+LoanRequiredBy+"','"+LoanReturned+"',"+LoanReferral+","+GrantReferral+","+LowUtilization+","+Interim+","+TechConsult+","+Session("insStaff_id")+","+OtherReferral+","+ConsultationOnly+",0,'A',0)}";
	rsPILATReferral.CursorType = 0;
	rsPILATReferral.CursorLocation = 2;
	rsPILATReferral.LockType = 3;
	rsPILATReferral.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New PILAT Referral</title>
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
				document.frm0201.reset();
			break;			
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Init() {
		document.frm0201.ReferralDate.focus();
	}
		
	function Save(){
		if ((!CheckDate(document.frm0201.ReferralDate.value)) && (document.frm0201.ReferralDate.value != "")){
			alert("Invalid Referral Date.");
			document.frm0201.ReferralDate.focus();
			return ;
		}
		if (!CheckDate(document.frm0201.LoanRequiredBy.value)) {
			alert("Invalid Loan Required By Date.");
			document.frm0201.LoanRequiredBy.focus();
			return ;
		}
		if (!CheckDate(document.frm0201.LoanReturned.value)) {
			alert("Invalid Loan Returned By Date.");
			document.frm0201.LoanReturned.focus();
			return ;
		}
		document.frm0201.submit();
	}
	
	function ChangeReferralType(){
		if (document.frm0201.ReferralType[3].checked) {
			InterimDetails.style.visibility = "visible";		
		} else {
			InterimDetails.style.visibility = "hidden";
			document.frm0201.HardwareSoftwarePurchasedByInstitution.checked = false;
			document.frm0201.PurchaseOrderSigned.checked = false;
			document.frm0201.LoanReturned.value = "";
			document.frm0201.LoanRequiredBy.value = "";			
		}	
	}	
	</script>
</head>
<body onLoad="Init();"> 
<form name="frm0201" method="POST" action="<%=MM_insertAction%>">
<h5>New PILAT Referral</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Date:</td>
		<td nowrap>
			<input type="text" name="ReferralDate" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Type:</td>
		<td width="130" class="blue">Loan</td>
		<td width="130" class="blue">Buyout</td>
		<td width="130" class="blue">Other</td>
	</tr>
    <tr> 
		<td></td>
		<td nowrap><input type="radio" name="ReferralType" onClick="ChangeReferralType();" value="1" tabindex="2" CHECKED class="chkstyle">Low Utilization</td>
		<td nowrap><input type="radio" name="ReferralType" onClick="ChangeReferralType();" value="3" tabindex="4" class="chkstyle">Tech Consult/Purchase</td>
		<td nowrap><input type="radio" name="ReferralType" onClick="ChangeReferralType();" value="4" tabindex="5" accesskey="L" class="chkstyle">Consultation Only</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="radio" name="ReferralType" onClick="ChangeReferralType();" value="2" tabindex="3" class="chkstyle">Interim</td>
		<td></td>
		<td></td>		
    </tr>
</table>
<div id="InterimDetails" style="visibility: hidden">
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap colspan="4">
			Hardware/software to be purchased or replaced by institution:&nbsp;
			<input type="checkbox" name="HardwareSoftwarePurchasedByInstitution" tabindex="6" class="chkstyle">
		</td>
	</tr>
	<tr>
		<td nowrap colspan="4">
			Purchase order been signed:&nbsp;
			<input type="checkbox" name="PurchaseOrderSigned" tabindex="7" class="chkstyle">
		</td>
	</tr>
	<tr>
		<td nowrap colspan="2">Equipment Loan is required by:</td>
		<td nowrap colspan="2">
			<input type="textbox" name="LoanRequiredBy" maxlength="10" size="11" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap colspan="2">Equipment Loan to be returned by:</td>
		<td nowrap colspan="2">
			<input type="textbox" name="LoanReturned" maxlength="10" size="11" tabindex="9" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>		
	</tr>
</table>
</div>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="10" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="top.window.close();" tabindex="11" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>