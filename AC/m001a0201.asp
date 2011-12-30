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
	var EPPDPostSecondaryLoan = ((Request.Form("EPPDPostSecondaryLoan")=="1")?"1":"0");	
	var EPPDEmployLoan = ((Request.Form("EPPDEmployLoan")=="1")?"1":"0");		
	var EPPDPSTPLoan = ((Request.Form("EPPDPSTPLoan")=="1")?"1":"0");	
	var EPPDTrainingLoan = ((Request.Form("EPPDTrainingLoan")=="1")?"1":"0");							
	var PostSecondaryCSG = ((Request.Form("PostSecondaryCSG")=="1")?"1":"0");	
	var PostSecondaryAPSD = ((Request.Form("PostSecondaryAPSD")=="1")?"1":"0");	
	var EPPDCSG = ((Request.Form("EPPDCSG")=="1")?"1":"0");								

	var LoanReferral = (((EPPDPostSecondaryLoan=="1") || (EPPDEmployLoan=="1") || (EPPDPSTPLoan=="1") || (EPPDTrainingLoan=="1"))?"1":"0");	
	var GrantReferral = (((PostSecondaryCSG=="1") || (PostSecondaryAPSD=="1") || (EPPDCSG=="1"))?"1":"0");
	
	var rsReferral = Server.CreateObject("ADODB.Recordset");
	rsReferral.ActiveConnection = MM_cnnASP02_STRING;
	rsReferral.Source = "{call dbo.cp_referrals2("+Request.QueryString("intAdult_id")+",0,0,'"+Request.Form("ReferralDate")+"',"+LoanReferral+","+GrantReferral+","+EPPDPostSecondaryLoan+","+EPPDEmployLoan+","+EPPDPSTPLoan+","+EPPDTrainingLoan+","+PostSecondaryCSG+","+PostSecondaryAPSD+","+EPPDCSG+",0,'A',0)}";
	rsReferral.CursorType = 0;
	rsReferral.CursorLocation = 2;
	rsReferral.LockType = 3;
	rsReferral.Open();
	Response.Redirect("InsertSuccessful.html");
}

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client("+ Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();
%>
<html>
<head>
	<title>New Referral Record</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/Myfunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function Save(){
		if (!CheckDate(document.frm0201.ReferralDate.value)){
			alert("Invalid Date.");
			document.frm0201.ReferralDate.focus();
			return ;
		}
		document.frm0201.submit();
	}
	</script>	
</head>
<body onLoad="javascript:document.frm0201.ReferralDate.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0201">
<h5>New Referral for <%=(rsClient.Fields.Item("chvName").Value)%>:</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Date:</td>
		<td nowrap>
			<input type="text" name="ReferralDate" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap height="20">Referral Type:</td>
		<td align="center" class="headrow" width="150">Loan</td>
		<td align="center" class="headrow" width="150">Grant</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="EPPDPostSecondaryLoan" value="1" tabindex="2" class="chkstyle">EPPD-PS-Loan</td>
		<td nowrap><input type="checkbox" name="PostSecondaryCSG" value="1" tabindex="6" class="chkstyle" >PS-CSG</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="EPPDEmployLoan" value="1" tabindex="3" class="chkstyle">EPPD-Employ-Loan</td>
		<td nowrap><input type="checkbox" name="PostSecondaryAPSD" value="1" tabindex="7" class="chkstyle">PS-APSD</td>
    </tr>
    <tr>
		<td></td>
		<td nowrap><input type="checkbox" name="EPPDPSTPLoan" value="1" tabindex="4" class="chkstyle">EPPD-PSTP-Loan</td>
		<td nowrap><input type="checkbox" name="EPPDCSG" value="1" tabindex="8" class="chkstyle">EPPD-CSG</td>
    </tr>
    <tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="EPPDTrainingLoan" value="1" tabindex="5" class="chkstyle">EPPD-Training-Loan</td>
		<td></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="9" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="10" onClick="window.close();" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsClient.Close();
%>