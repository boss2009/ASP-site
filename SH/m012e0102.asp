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
	var ChkFundingSource = Server.CreateObject("ADODB.Command");
	ChkFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	ChkFundingSource.CommandText = "dbo.cp_Chk_School_FundSrc";
	ChkFundingSource.CommandType = 4;
	ChkFundingSource.CommandTimeout = 0;
	ChkFundingSource.Prepared = true;
	ChkFundingSource.Parameters.Append(ChkFundingSource.CreateParameter("RETURN_VALUE", 3, 4));
	ChkFundingSource.Parameters.Append(ChkFundingSource.CreateParameter("@insSchool_id", 3, 1,10000,Request.QueryString("insSchool_id")));
	ChkFundingSource.Parameters.Append(ChkFundingSource.CreateParameter("@intRfral_id", 3, 1,10000,Request.Form("MM_recordId")));	
	ChkFundingSource.Parameters.Append(ChkFundingSource.CreateParameter("@insRef_Agt_id", 3, 1,10000,Request.Form("MM_RefAgentId")));		
	ChkFundingSource.Parameters.Append(ChkFundingSource.CreateParameter("@insRtnFlag", 2, 2));
	ChkFundingSource.Execute();
	
	//if funding source exist, do update, else add.
	var rsInstitutionFundingSource = Server.CreateObject("ADODB.Recordset");
	rsInstitutionFundingSource.ActiveConnection = MM_cnnASP02_STRING;
	if (ChkFundingSource.Parameters.Item("@insRtnFlag").Value=="1") {
		rsInstitutionFundingSource.Source = "{call dbo.cp_school_funding_src("+Request.Form("intScool_FundSrc_id")+","+Request.QueryString("insSchool_id")+","+Request.Form("MM_recordId")+","+Request.Form("MM_RefAgentId")+","+Request.Form("FundingSource")+",0,'E',0)}";
	} else {
		rsInstitutionFundingSource.Source = "{call dbo.cp_school_funding_src(0,"+Request.QueryString("insSchool_id")+","+Request.Form("MM_recordId")+","+Request.Form("MM_RefAgentId")+","+Request.Form("FundingSource")+",0,'A',0)}";
	}
	rsInstitutionFundingSource.CursorType = 0;
	rsInstitutionFundingSource.CursorLocation = 2;
	rsInstitutionFundingSource.LockType = 3;
	rsInstitutionFundingSource.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012q0102.asp&insSchool_id="+Request.QueryString("insSchool_id"));
}

var rsInstitutionFundingSource = Server.CreateObject("ADODB.Recordset");
rsInstitutionFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsInstitutionFundingSource.Source = "{call dbo.cp_school_funding_src(0,"+Request.QueryString("insSchool_id")+","+Request.QueryString("intReferral_id")+","+Request.QueryString("insRefAgt_id")+",0,1,'Q',0)}";
rsInstitutionFundingSource.CursorType = 0;
rsInstitutionFundingSource.CursorLocation = 2;
rsInstitutionFundingSource.LockType = 3;
rsInstitutionFundingSource.Open();

var refagt = ((!rsInstitutionFundingSource.EOF)?rsInstitutionFundingSource.Fields.Item("insRefAgt_id").Value:0);

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_query_refagt_fundsrc("+refagt+",0)}";
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
				document.frm0102.submit();
			break;
			case 85:
				//alert("U");
				document.frm0102.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
			break;
		}
	}
	</script>	
</head>
<body onLoad="document.frm0102.FundingSource.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0102">
<h5>Funding Source</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Referral ID:</td>
		<td nowrap><%=ZeroPadFormat(rsInstitutionFundingSource.Fields.Item("intReferral_id").Value,8)%></td>
    </tr>
    <tr> 
		<td nowrap>Referral Date:</td>
		<td nowrap><%=FilterDate(rsInstitutionFundingSource.Fields.Item("dtsRefral_date").Value)%></td>
	</tr>
	<tr>
		<td nowrap>Referral Type:</td>
		<td nowrap><%=(rsInstitutionFundingSource.Fields.Item("chvRefAgt").Value)%></td>
    </tr>
    <tr> 
		<td nowrap>Funding Source:</td>
		<td nowrap><select name="FundingSource" tabindex="1">
		<% 
		while (!rsFundingSource.EOF) {
		%>
			<option value="<%=(rsFundingSource.Fields.Item("insFunding_source_id").Value)%>" <%=((rsFundingSource.Fields.Item("insFunding_source_id").Value == rsInstitutionFundingSource.Fields.Item("insSel_Funding_Source").Value)?"SELECTED":"")%> ><%=(rsFundingSource.Fields.Item("chvfunding_source_name").Value)%></option>
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
<input type="hidden" name="MM_RefAgentId" value="<%=rsInstitutionFundingSource.Fields.Item("insRefAgt_id").Value%>">
<input type="hidden" name="MM_recordId" value="<%=rsInstitutionFundingSource.Fields.Item("intReferral_id").Value%>">
<input type="hidden" name="intScool_FundSrc_id" value="<%=rsInstitutionFundingSource.Fields.Item("intScool_FundSrc_id").Value%>">
</form>
</body>
</html>
<%
rsFundingSource.Close();
rsInstitutionFundingSource.Close();
%>