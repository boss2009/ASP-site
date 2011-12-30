<!--
This page is current not being used.  When a referral (grant) has been added for the client,
a new grant eligibility is created automatically.
-->
<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true"){
	var bitIs_PS_CSG = ((Request.Form("PostSecondaryCSG")=="1") ? "1":"0");
	var bitIs_PS_APSD = ((Request.Form("PostSecondaryAPSD")=="1") ? "1":"0");
	var bitIs_EPPD_CSG = ((Request.Form("EPPDCSG")=="1") ? "1":"0");
	var rsGrant = Server.CreateObject("ADODB.Recordset");
	rsGrant.ActiveConnection = MM_cnnASP02_STRING;
	rsGrant.Source = "{call dbo.cp_grant_elgbty2(0,"+ Request.QueryString("intAdult_id") + ",'"+Request.Form("ReferralDate")+"','"+Request.Form("EligibleFrom")+"','"+Request.Form("EligibleTo")+"',"+ bitIs_PS_CSG+","+ bitIs_PS_APSD+","+ bitIs_EPPD_CSG+","+Request.Form("GrantAmount")+",0,0,'A',0)}";
	rsGrant.CursorType = 0;
	rsGrant.CursorLocation = 2;
	rsGrant.LockType = 3;
	rsGrant.Open();
	Response.Redirect("InsertSuccessful.html");
}

var ChkGrantEligibility = Server.CreateObject("ADODB.Command");
ChkGrantEligibility.ActiveConnection = MM_cnnASP02_STRING;
ChkGrantEligibility.CommandText = "dbo.cp_Chk_Grant_Elgbty";
ChkGrantEligibility.CommandType = 4;
ChkGrantEligibility.CommandTimeout = 0;
ChkGrantEligibility.Prepared = true;
ChkGrantEligibility.Parameters.Append(ChkGrantEligibility.CreateParameter("RETURN_VALUE", 3, 4));
ChkGrantEligibility.Parameters.Append(ChkGrantEligibility.CreateParameter("@intAdult_id", 3, 1,1000,Request.QueryString("intAdult_id")));
ChkGrantEligibility.Parameters.Append(ChkGrantEligibility.CreateParameter("@insRtnFlag", 2, 2));
ChkGrantEligibility.Execute();
%>
<html>
<head>
	<title>New Grant Eligibility</title>
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
				document.frm0603.reset();
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
		if (isNaN(document.frm0603.GrantAmount.value)){
			alert("Invalid Grant Amount.");
			document.frm0603.GrantAmount.focus();
			return ;
		}
		if (!CheckDate(document.frm0603.EligibleFrom.value)){
			alert("Invalid Eligible From Date.");
			document.frm0603.EligibleFrom.focus();
			return ;
		}
		if (!CheckDate(document.frm0603.EligibleTo.value)){
			alert("Invalid Eligible To Date.");
			document.frm0603.EligibleTo.focus();
			return ;		
		}
		var count = 0
		if (document.frm0603.PostSecondaryCSG.checked==true) count++;
		if (document.frm0603.PostSecondaryAPSD.checked==true) count++;
		if (document.frm0603.EPPDCSG.checked==true) count++;
		if (count < 1) {
			alert("Select at least one grant type.");
			return;
		}
		document.frm0603.submit();
		document.frm0603.btnSave.disabled = true;
	}
	</script>
</head>
<body onLoad="javascript:document.frm0603.ReferralDate.focus()">
<form name="frm0603" method="POST" action="<%=MM_editAction%>">
<h5>New Grant Eligibility</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Referral Date:</td>
		<td nowrap>
			<input type="text" name="ReferralDate" value="<%=FilterDate(CurrentDate())%>"  size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Referral Type:</td>
		<td nowrap><select name="ReferralType" tabindex="3" disabled>
			<option value="0" <%=((ChkGrantEligibility.Parameters.Item("@insRtnFlag").Value == 0)?"SELECTED":"")%>>Referral
			<option value="1" <%=((ChkGrantEligibility.Parameters.Item("@insRtnFlag").Value == 1)?"SELECTED":"")%>>Re-referral
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Grant Type:</td>
		<td nowrap>
	  		<input type="checkbox" name="PostSecondaryCSG" value="1" tabindex="4" class="chkstyle">PS-CSG
			<input type="checkbox" name="PostSecondaryAPSD" value="1" tabindex="5" class="chkstyle">PS-APSD
			<input type="checkbox" name="EPPDCSG" value="1" tabindex="6" class="chkstyle">EPPD-CSG
		</td>
    </tr>
    <tr> 
		<td nowrap>Grant Amount:</td>
		<td nowrap>$<input type="text" name="GrantAmount" value="8000.00" size="6" tabindex="7" onKeypress="AllowNumericOnly();"></td>
    </tr>
    <tr> 
		<td nowrap>Eligibility</td>
		<td></td>
	</tr>
	<tr>
		<td nowrap align="right">From:</td>
		<td nowrap> 
			<input type="text" name="EligibleFrom" value="" maxlength="10" size="11" tabindex="8" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap align="right">To:</td>
		<td nowrap>
			<input type="text" name="EligibleTo" value="" maxlength="10" size="11" tabindex="10" accesskey="L" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" name="btnSave" value="Save" tabindex="11" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="12" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>