<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true"){
	var rsCheckSelected = Server.CreateObject("ADODB.Recordset");
	rsCheckSelected.ActiveConnection = MM_cnnASP02_STRING;
	rsCheckSelected.Source = "{call dbo.cp_buyout_funding_src("+Request.QueryString("intBuyout_req_id")+","+Request.Form("MM_recordId")+",0,0,1,'Q',0)}";
	rsCheckSelected.CursorType = 0;
	rsCheckSelected.CursorLocation = 2;
	rsCheckSelected.LockType = 3;
	rsCheckSelected.Open();

	var rsBuyoutFundingSource = Server.CreateObject("ADODB.Recordset");
	rsBuyoutFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	rsBuyoutFundingSource.CursorType = 0;
	rsBuyoutFundingSource.CursorLocation = 2;
	rsBuyoutFundingSource.LockType = 3;	
	if (rsCheckSelected.Fields.Item("bitIs_Sel_FundingSrc").Value=="1") {
		rsBuyoutFundingSource.Source = "{call dbo.cp_buyout_funding_src("+Request.QueryString("intBuyout_req_id")+","+Request.Form("MM_recordId")+","+Request.Form("MM_RefAgentId")+","+Request.Form("FundingSource")+",0,'E',0)}";	
		rsBuyoutFundingSource.Open();	
	} else {
		rsBuyoutFundingSource.Source = "{call dbo.cp_buyout_funding_src("+Request.QueryString("intBuyout_req_id")+","+Request.Form("MM_recordId")+","+Request.Form("MM_RefAgentId")+","+Request.Form("FundingSource")+",0,'A',0)}";	
		rsBuyoutFundingSource.Open();		
	}
	Response.Redirect("UpdateSuccessful.asp?page=m010q0103.asp&intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));
}

var rsBuyoutFundingSource = Server.CreateObject("ADODB.Recordset");
rsBuyoutFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsBuyoutFundingSource.Source = "{call dbo.cp_buyout_funding_src("+Request.QueryString("intBuyout_req_id")+","+Request.QueryString("intReferral_id")+",0,0,1,'Q',0)}";
rsBuyoutFundingSource.CursorType = 0;
rsBuyoutFundingSource.CursorLocation = 2;
rsBuyoutFundingSource.LockType = 3;
rsBuyoutFundingSource.Open();

var ReferralAgent = ((!rsBuyoutFundingSource.EOF)?rsBuyoutFundingSource.Fields.Item("insRefAgt_id").Value:0);

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_query_refagt_fundsrc("+ReferralAgent+",0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();
%>
<html>
<head>
	<title>Update Funding Source</title>
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
			case 85:
				//alert("U");
				document.frm0103.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
</head>
<body onLoad="document.frm0103.FundingSource.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0103">
<h5>Update Funding Source</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral ID:</td>
		<td nowrap><%=ZeroPadFormat(rsBuyoutFundingSource.Fields.Item("intReferral_id").Value,8)%></td>
    </tr>
    <tr> 
		<td nowrap>Referral Date:</td>
		<td nowrap><%=FilterDate(rsBuyoutFundingSource.Fields.Item("dtsRefral_date").Value)%></td>
	</tr>
	<tr>
		<td nowrap>Referral Type:</td>
		<td nowrap><%=(rsBuyoutFundingSource.Fields.Item("chvRefAgt").Value)%></td>
    </tr>
    <tr> 
		<td nowrap>Funding Source:</td>
		<td nowrap><select name="FundingSource" tabindex="1">
		<% 
		while (!rsFundingSource.EOF) {
		%>
			<option value="<%=(rsFundingSource.Fields.Item("insFunding_source_id").Value)%>" <%=((rsFundingSource.Fields.Item("insFunding_source_id").Value == rsBuyoutFundingSource.Fields.Item("insSel_Funding_Source").Value)?"SELECTED":"")%> ><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%></option>
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
		<td><input type="submit" value="Save" tabindex="2" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_RefAgentId" value="<%=rsBuyoutFundingSource.Fields.Item("insRefAgt_id").Value%>">
<input type="hidden" name="MM_recordId" value="<%=rsBuyoutFundingSource.Fields.Item("intReferral_id").Value%>">
</form>
</body>
</html>
<%
rsFundingSource.Close();
rsBuyoutFundingSource.Close();
%>