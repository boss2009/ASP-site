<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
// set the form action variable
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {	
	var ReferringAgentName = String(Request.Form("ReferringAgentName")).replace(/'/g, "''");	
	var IsActive = ((Request.Form("IsActive")=="1") ? "1":"0");
	var IsLoan = ((Request.Form("IsLoan")=="1") ? "1":"0");
	var IsBuyout = ((Request.Form("IsBuyout")=="1") ? "1":"0");		
	var rsReferringAgent = Server.CreateObject("ADODB.Recordset");
	rsReferringAgent.ActiveConnection = MM_cnnASP02_STRING;
	rsReferringAgent.Source = "{call dbo.cp_update_referring_agent("+ Request.Form("MM_recordId") + ",'" + Request.Form("ReferringAgentName") + "'," + IsActive + "," + IsLoan + "," + IsBuyout + ",'" + Request.Form("FundingSourceCode") + "',0)}";
	rsReferringAgent.CursorType = 0;
	rsReferringAgent.CursorLocation = 2;
	rsReferringAgent.LockType = 3;
	rsReferringAgent.Open();
	Response.Redirect("m018q0312.asp");
}
%>
<html>
<head>
	<title>Update Referring Agent</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
</head>
<body>
<form name="frm0312b" method="POST" action="<%=MM_editAction%>">
<h5>Update Referring Agent</h5>
<i><b>This action may affect the integrity of the the system.</b><br>
To confirm the change, click [Proceed].  To Abort, click [Cancel].</i>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Referring Agent Name:</td>
		<td nowrap><input type="text" name="ReferringAgentName" value="<%=(Request.QueryString("chvname"))%>" readonly accesskey="F"></td>
    </tr>
    <tr> 
		<td nowrap>Is Active</td>
		<td nowrap><input type="text" name="IsActive" value="<%=(Request.QueryString("bitis_active"))%>" readonly></td>
    </tr>
    <tr> 
		<td nowrap>Is Loan:</td>
		<td nowrap><input type="text" name="IsLoan" value="<%=(Request.QueryString("bitis_loan"))%>" readonly></td>
    </tr>
    <tr> 
		<td nowrap>Is Buyout:</td>
		<td nowrap><input type="text" name="IsBuyOut" value="<%=(Request.QueryString("bitis_BuyOut"))%>" readonly></td>
    </tr>
    <tr> 
		<td nowrap>Funding Source:</td>
		<td nowrap><input type="text" name="FundingSourceCode" value="<%=(Request.QueryString("chrFS_chbx"))%>" readonly accesskey="L"></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="submit" value="Proceed" class="btnstyle"></td>
		<td><input type="button" value="Cancel" onClick="history.go(-2)" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=Request.QueryString("insrefer_agent_id")%>">
</form>
</body>
</html>