<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true"){
	var rsBuyoutFundingSource = Server.CreateObject("ADODB.Recordset");
	rsBuyoutFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutFundingSource.Source = "{call dbo.cp_buyout_funding_src("+Request.QueryString("intBuyout_req_id")+","+Request.Form("ReferralDate")+","+Request.Form("ReferralType")+","+Request.Form("FundingSource")+",0,'A',0)}";
	rsBuyoutFundingSource.CursorType = 0;
	rsBuyoutFundingSource.CursorLocation = 2;
	rsBuyoutFundingSource.LockType = 3;
	rsBuyoutFundingSource.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsBuyout = Server.CreateObject("ADODB.Recordset");
rsBuyout.ActiveConnection = MM_cnnASP02_STRING;
rsBuyout.Source = "{call dbo.cp_Buyout_request3("+ Request.QueryString("intBuyout_Req_id") + ",0,0,'',0,'',0,0,0,0,1,'Q',0)}";
rsBuyout.CursorType = 0;
rsBuyout.CursorLocation = 2;
rsBuyout.LockType = 3;
rsBuyout.Open();

var IsClient = true;
if (!rsBuyout.EOF) {
	if ((rsBuyout.Fields.Item("insEq_user_type").Value == 3) && (rsBuyout.Fields.Item("intEq_user_id").Value > 0)) {
		var rsReferral = Server.CreateObject("ADODB.Recordset");
		rsReferral.ActiveConnection = MM_cnnASP02_STRING;
		rsReferral.Source = "{call dbo.cp_referrals2("+rsBuyout.Fields.Item("intEq_user_id").Value+",0,0,'',0,1,0,0,0,0,0,0,0,4,'Q',0)}";
		rsReferral.CursorType = 0;
		rsReferral.CursorLocation = 2;
		rsReferral.LockType = 3;
		rsReferral.Open();

		var refid = ((String(Request.Form("ReferralDate"))!="undefined")?Request.Form("ReferralDate"):rsReferral.Fields.Item("intReferral_id").value);
		var rsReferralType = Server.CreateObject("ADODB.Recordset");
		rsReferralType.ActiveConnection = MM_cnnASP02_STRING;
		rsReferralType.Source = "{call dbo.cp_asp_lkup2(11,"+refid+",'',0,'',0)}";
		rsReferralType.CursorType = 0;
		rsReferralType.CursorLocation = 2;
		rsReferralType.LockType = 3;
		rsReferralType.Open();		

		var rsFundingSource = Server.CreateObject("ADODB.Recordset");
		rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
		rsFundingSource.Source = "{call dbo.cp_funding_source_attributes(0,0,1,0,0,0,1,0,2,'Q',0)}";
		rsFundingSource.CursorType = 0;
		rsFundingSource.CursorLocation = 2;
		rsFundingSource.LockType = 3;
		rsFundingSource.Open();		
	} else {
		IsClient = false;
	}
} else {
	IsClient = false;
}
%>
<html>
<head>
	<title>New Funding Source</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83:
				//alert("S");
				document.frm0103.submit();
			break;
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	function ChangeReferral(){
		document.frm0103.MM_insert.value = "false";
		document.frm0103.submit();
	}
	
	function Save(){ 
 		if (document.frm0103.ReferralDate.value < 1) {
			alert("Select a Referral Date.");
			return; 
		}
		if (document.frm0103.ReferralType.value < 1) {
			alert("Select a Referral Type.");
			return; 
		}
		if (document.frm0103.FundingSource.value < 1) {
			alert("Select a Funding Source.");
			return; 
		}
		document.frm0103.MM_insert.value = "true";		
		document.frm0103.submit();
	}
	</script>
</head>
<body onLoad="document.frm0103.ReferralDate.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0103">
<h5>New Funding Source</h5>
<%
if (!IsClient) {
%>
Information not available.  Either the buyer is institutional or the client is not found.
<%
} else {
%>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral Date:</td>
		<td nowrap><select name="ReferralDate" tabindex="1" accesskey="F" onChange="ChangeReferral();">
			<% 
			while (!rsReferral.EOF) {
			%>
				<option value="<%=(rsReferral.Fields.Item("intReferral_id").Value)%>" <%=((rsReferral.Fields.Item("intReferral_id").Value==Request.Form("ReferralDate"))?"SELECTED":"")%>><%=(rsReferral.Fields.Item("dtsRefral_date").Value)%></option>
			<%
				rsReferral.MoveNext();
			}
			%>			
		</select></td>
    </tr>
	<tr>
		<td nowrap>Referral Type:</td>
		<td nowrap><select name="ReferralType" tabindex="2">
			<% 
			while (!rsReferralType.EOF) {
			%>
				<option value="<%=(rsReferralType.Fields.Item("insRefAgt_id").Value)%>"><%=(rsReferralType.Fields.Item("chvReferring_Agent").Value)%>
			<%
				rsReferralType.MoveNext();
			}
			%>
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Funding Source:</td>
		<td nowrap><select name="FundingSource" tabindex="3">
			<% 
			while (!rsFundingSource.EOF) {
			%>
				<option value="<%=(rsFundingSource.Fields.Item("insFunding_source_id").Value)%>"><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%></option>
			<%
				rsFundingSource.MoveNext();
			}
			%>
		</select></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="4" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="5" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<%
}
%>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>