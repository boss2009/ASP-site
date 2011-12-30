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
	rsReferral.Source = "{call dbo.cp_referrals2("+Request.QueryString("intAdult_id")+","+Request.QueryString("intReferral_id")+",0,'"+Request.Form("ReferralDate")+"',"+LoanReferral+","+GrantReferral+","+EPPDPostSecondaryLoan+","+EPPDEmployLoan+","+EPPDPSTPLoan+","+EPPDTrainingLoan+","+PostSecondaryCSG+","+PostSecondaryAPSD+","+EPPDCSG+",0,'E',0)}";
	rsReferral.CursorType = 0;
	rsReferral.CursorLocation = 2;
	rsReferral.LockType = 3;
	rsReferral.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m001q0201.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsReferral = Server.CreateObject("ADODB.Recordset");
rsReferral.ActiveConnection = MM_cnnASP02_STRING;
rsReferral.Source = "{call dbo.cp_referrals2(0,"+ Request.QueryString("intReferral_id") + ",0,'',0,0,0,0,0,0,0,0,0,1,'Q',0)}";
rsReferral.CursorType = 0;
rsReferral.CursorLocation = 2;
rsReferral.LockType = 3;
rsReferral.Open();
%>
<html>
<head>
	<title>Update Referral</title>
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
			case 85:
				//alert("U");
				document.frm0201.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
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
<h5>Update Referral</h5>
<hr>
<table cellpadding="1" cellspacing="1">
<!--
	<tr>
		<td>Referral Class:</td>
		<td><select name="ReferralClass" tabindex="1" accesskey="F">
			<option value=0 <%=((rsReferral.Fields.Item("bitIs_New").Value == 0)?"SELECTED":"")%>>Referral
			<option value=1 <%=((rsReferral.Fields.Item("bitIs_New").Value == 1)?"SELECTED":"")%>>Re-Referral			
		</select></td> 
	</tr>
-->
	<tr>
		<td nowrap>Referral Date:</td>
		<td nowrap>
			<input type="text" name="ReferralDate" value="<%=FilterDate(rsReferral.Fields.Item("dtsRefral_date").Value)%>" maxlength="10" tabindex="2" size="11" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap height="20">Referral Type:</td>
		<td nowrap align="center" width="150" class="headrow">Loan</td>
		<td nowrap align="center" width="150" class="headrow">Grant</td>
    </tr>
	<tr> 
		<td></td>
		<td nowrap><input type="checkbox" name="EPPDPostSecondaryLoan" <%=((rsReferral.Fields.Item("bitIs_PS_Loan").Value == 1)?"CHECKED":"")%> value="1" tabindex="5" class="chkstyle">EPPD-PS-Loan</td>
		<td nowrap><input type="checkbox" name="PostSecondaryCSG" <%=((rsReferral.Fields.Item("bitIs_PS_CSG_Grant").Value == 1)?"CHECKED":"")%> value="1" tabindex="9" class="chkstyle">PS-CSG</td>
	</tr>
	<tr> 
		<td></td>	
		<td nowrap><input type="checkbox" name="EPPDEmployLoan" <%=((rsReferral.Fields.Item("bitIs_VRS_Emply_Loan").Value == 1)?"CHECKED":"")%> value="1" tabindex="6" class="chkstyle">EPPD-Employ-Loan</td>
		<td nowrap><input type="checkbox" name="PostSecondaryAPSD" <%=((rsReferral.Fields.Item("bitIs_PS_APSD_Grant").Value == 1)?"CHECKED":"")%> value="1" tabindex="10" class="chkstyle">PS-APSD</td>
	</tr>
	<tr> 
		<td></td>	
		<td nowrap><input type="checkbox" name="EPPDPSTPLoan" <%=((rsReferral.Fields.Item("bitIs_VRS_PSTP_Loan").Value == 1)?"CHECKED":"")%> value="1" tabindex="7" class="chkstyle">EPPD-PSTP-Loan</td>
		<td nowrap><input type="checkbox" name="EPPDCSG" <%=((rsReferral.Fields.Item("bitIs_VRS_CSG_Grant").Value == 1)?"CHECKED":"")%> value="1" tabindex="11" class="chkstyle">EPPD-CSG</td>
	</tr>
    <tr>
		<td></td>	
		<td nowrap><input type="checkbox" name="EPPDTrainingLoan" <%=((rsReferral.Fields.Item("bitIs_VRS_Train_Loan").Value == 1)?"CHECKED":"")%> value="1" tabindex="8" class="chkstyle">EPPD-Training-Loan</td>
		<td></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="12" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="13" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="history.back();" tabindex="14" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsReferral.Fields.Item("intReferral_id").Value %>">
</form>
</body>
</html>
<%
rsReferral.Close();
%>