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
	var AmountDue = ((Request.Form("AmountDue")=="")?"0":Request.Form("AmountDue"));
	var AmountReceived = ((Request.Form("AmountDue")=="")?"0":Request.Form("AmountReceived"));	
	var PaidByEmployer = ((Request.Form("PaidByEmployer")=="1")?"1":"0");	
	var PaidByClient = ((Request.Form("PaidByClient")=="1")?"1":"0");	
	var PaidByEPPDConsultant = ((Request.Form("PaidByEPPDConsultant")=="1")?"1":"0");			
	var rsFollowUp = Server.CreateObject("ADODB.Recordset");
	rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
	rsFollowUp.Source="{call dbo.cp_follow_up(0,'2',"+ Request.QueryString("intAdult_id") +",0,'',0,0,0,'','',0,0,0,0,'',0,'"+Request.Form("BuyoutDueDate")+"','"+Request.Form("InvoiceNumber")+"','"+Request.Form("CaseNumber")+"',"+Request.Form("Case")+",'"+Request.Form("InvoiceSentDate")+"',"+AmountDue+","+AmountReceived+",'"+Request.Form("DateReceived")+"',"+PaidByEmployer+","+PaidByClient+","+PaidByEPPDConsultant+",'"+Request.Form("DefaultedDate")+"','"+Request.Form("DateCleared")+"','',0,'','',0,'A',0)}";
	rsFollowUp.CursorType = 0;
	rsFollowUp.CursorLocation = 2;
	rsFollowUp.LockType = 3;
	rsFollowUp.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New EPPD Buyout Follow-Up</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="javascript" src="../js/MyFunctions.js"></script>
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
	<script language="javascript">
	function Save(){
		if (!CheckDate(document.frm1102.BuyoutDueDate.value)) {
			alert("Invalid Buyout Due Date.");
			document.frm1102.BuyoutDueDate.focus();
			return ;
		}
		if (!CheckDate(document.frm1102.InvoiceSentDate.value)) {
			alert("Invalid Invoice Sent Date.");
			document.frm1102.InvoiceSentDate.focus();
			return ;
		}
		if (isNaN(document.frm1102.AmountDue.value)) {
			alert("Invalid Amount Due.");
			document.frm1102.AmountDue.focus();
			return ;
		}
		if (isNaN(document.frm1102.AmountReceived.value)) {
			alert("Invalid Amount Received.");
			document.frm1102.AmountReceived.focus();
			return ;
		}
		if (!CheckDate(document.frm1102.DateReceived.value)) {
			alert("Invalid Payment Date Received.");
			document.frm1102.DateReceived.focus();
			return ;
		}
		if (!CheckDate(document.frm1102.DefaultedDate.value)) {
			alert("Invalid Defaulted date.");
			document.frm1102.DefaultedDate.focus();
			return ;
		}
		if (!CheckDate(document.frm1102.DateCleared.value)) {
			alert("Invalid Date Cleared.");
			document.frm1102.DateCleared.focus();
			return ;
		}
		document.frm1102.submit();
	}
	</script>
</head>
<body onLoad="document.frm1102.BuyoutDueDate.focus();">
<form name="frm1102" method="POST" action="<%=MM_editAction%>">
<h5>New EPPD Buyout Follow-Up</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Buyout Due Date:</td>
		<td nowrap>
			<input type="text" name="BuyoutDueDate" size="11" maxlength="10" tabindex="1" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
	</tr>
		<td nowrap>Invoice Number:</td>
		<td nowrap><input type="text" name="InvoiceNumber" maxlength="15" tabindex="2"></td>
	</tr>
	<tr> 
		<td nowrap>Invoice Sent Date:</td>
		<td nowrap>
			<input type="text" name="InvoiceSentDate" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
	</tr>
	<tr>
		<td nowrap>Case:</td>
		<td nowrap>
			<select name="Case" tabindex="4">
				<option value="1" SELECTED>New
				<option value="0">Outstanding
			</select>
			Number:
			<input type="text" name="CaseNumber" maxlength="15" tabindex="5" size="8">
		</td>
	</tr>
	<tr> 
		<td nowrap>Amount Due:</td>
		<td nowrap>$<input type="text" name="AmountDue" maxlength="8" size="8" tabindex="6"></td>
	</tr>
	<tr> 
		<td nowrap>Amount Received:</td>
		<td nowrap>$<input type="text" name="AmountReceived" maxlength="8" size="8" tabindex="7"></td>
	</tr>
	<tr>
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap>Paid in Full by:</td>
		<td nowrap>
			<input type="checkbox" name="PaidByEmployer" value="1" tabindex="9" class="chkstyle">Employer
			<input type="checkbox" name="PaidByClient" value="1" tabindex="10" class="chkstyle">Client
			<input type="checkbox" name="PaidByEPPDConsultant" value="1" tabindex="11" class="chkstyle">EPPD Consultant
		</td>
	</tr>
	<tr> 
		<td nowrap>Defaulted Date:</td>
		<td nowrap>
			<input type="text" name="DefaultedDate" size="11" maxlength="10" tabindex="12" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
	</tr>
	<tr>
		<td nowrap>Date Cleared:</td>
		<td nowrap>
			<input type="text" name="DateCleared" size="11" maxlength="10" tabindex="13" accesskey="L" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="14" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="15" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>