<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_updateAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_updateAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
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
	rsPILATReferral.Source = "{call dbo.cp_pilat_referrals3("+Request.QueryString("insSchool_id")+","+ Request.Form("ReferralID")+",'"+Request.Form("ReferralDate")+"',"+HardwareSoftwarePurchasedByInstitution+","+PurchaseOrderSigned+",'"+LoanRequiredBy+"','"+LoanReturned+"',"+LoanReferral+","+GrantReferral+","+LowUtilization+","+Interim+","+TechConsult+","+Session("insStaff_id")+","+OtherReferral+","+ConsultationOnly+",0,'E',0)}";
	rsPILATReferral.CursorType = 0;
	rsPILATReferral.CursorLocation = 2;
	rsPILATReferral.LockType = 3;
	rsPILATReferral.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012e0201.asp&intReferral_id="+Request.QueryString("intReferral_id")+"&insSchool_id="+Request.QueryString("insSchool_id"));
}

var rsPILATReferral = Server.CreateObject("ADODB.Recordset");
rsPILATReferral.ActiveConnection = MM_cnnASP02_STRING;
rsPILATReferral.Source = "{call dbo.cp_pilat_referrals3(0,"+Request.QueryString("intReferral_id")+",'',0,0,'','',0,0,0,0,0,0,0,0,1,'Q',0)}";
rsPILATReferral.CursorType = 0;
rsPILATReferral.CursorLocation = 2;
rsPILATReferral.LockType = 3;
rsPILATReferral.Open();	
%>
<html>
<head>
	<title>Update PILAT Referral</title>
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
				top.BodyFrame.location.href='m012q0201.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Init() {
		ChangeReferralType();
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
		var index=0;
		for (var i = 0; i < 4; i++) {
			if (document.frm0201.ReferralType[i].checked==1) index=i;
		}
		if (index==3) {
			InterimDetails.style.visibility = "visible";		
		} else {
			InterimDetails.style.visibility = "hidden";
			document.frm0201.HardwareSoftwarePurchasedByInstitution.checked=false;
			document.frm0201.PurchaseOrderSigned.checked=false;
			document.frm0201.LoanReturned.value="";
			document.frm0201.LoanRequiredBy.value="";
		}	
	}	</script>
</head>
<body onLoad="Init();"> 
<form action="<%=MM_updateAction%>" method="POST" name="frm0201">
<h5>Temp Referral</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Date:</td>
		<td nowrap>
			<input type="text" name="ReferralDate" value="<%=FilterDate(rsPILATReferral.Fields.Item("dtsRefral_date").Value)%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
</table>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Type:</td>
		<td width="130" class="headrow">Loan</td>
		<td width="130" class="headrow">Buyout</td>
		<td width="130" class="headrow">Other</td>		
	</tr>
    <tr> 
		<td></td>
		<td><input type="radio" name="ReferralType" <%=((rsPILATReferral.Fields.Item("bitIs_LowUtil_Loan").Value == 1)?"CHECKED":"")%> onClick="ChangeReferralType();" value="1" tabindex="2" class="chkstyle">Low Utilization</td>
		<td><input type="radio" name="ReferralType" <%=((rsPILATReferral.Fields.Item("bitIs_Consult_Grant").Value == 1)?"CHECKED":"")%> onClick="ChangeReferralType();" value="3" tabindex="4" class="chkstyle">Tech Consult/Purchase</td>
		<td><input type="radio" name="ReferralType" <%=((rsPILATReferral.Fields.Item("bitIs_Consult_Only").Value == 1)?"CHECKED":"")%> onClick="ChangeReferralType();" value="4" tabindex="5"class="chkstyle" accesskey="L" >Consultation Only</td>		
    </tr>
    <tr> 
		<td></td>
		<td><input type="radio" name="ReferralType" <%=((rsPILATReferral.Fields.Item("bitIs_Interim_Loan").Value == 1)?"CHECKED":"")%> onClick="ChangeReferralType();" value="2" tabindex="3" class="chkstyle">Interim</td>
		<td></td>
		<td></td>		
    </tr>
</table>
<div id="InterimDetails" style="visibility: hidden">
<table cellpadding="1" cellspacing="1">
	<tr>
		<td colspan="4">			
			<input type="checkbox" name="HardwareSoftwarePurchasedByInstitution" <%=((rsPILATReferral.Fields.Item("bitIs_RplBy_School").Value==1)?"CHECKED":"")%> tabindex="6" class="chkstyle">
			Hardware/software to be purchased or replaced by institution:&nbsp;
		</td>
	</tr>
	<tr>
		<td colspan="4">			
			<input type="checkbox" name="PurchaseOrderSigned" <%=((rsPILATReferral.Fields.Item("bitIs_Ord_Signed").Value==1)?"CHECKED":"")%> tabindex="7" class="chkstyle">
			Purchase order been signed:&nbsp;
		</td>
	</tr>
	<tr>
		<td colspan="2">Equipment Loan is required by:</td>
		<td colspan="2"><input type="textbox" name="LoanRequiredBy" value="<%=FilterDate(rsPILATReferral.Fields.Item("dtsRequired_by").Value)%>" maxlength="10" size="11" tabindex="8" onChange="FormatDate(this)"> <span style="font-size: 7pt">(mm/dd/yyyy)</span></td>
	</tr>
	<tr>
		<td colspan="2">Equipment Loan to be returned by:</td>
		<td colspan="2"><input type="textbox" name="LoanReturned" value="<%=FilterDate(rsPILATReferral.Fields.Item("dtsReturned_by").Value)%>" maxlength="10" size="11" tabindex="9" onChange="FormatDate(this)"> <span style="font-size: 7pt">(mm/dd/yyyy)</span></td>		
	</tr>
</table>
</div>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="10" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="11" class="btnstyle"></td>		
		<td><input type="button" value="Close" tabindex="12" onClick="top.BodyFrame.location.href='m012q0201.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>'" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="ReferralID" value="<%=Request.QueryString("intReferral_id")%>">
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsPILATReferral.Close();
%>