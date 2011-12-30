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
	var AmountDue = ((Request.Form("AmountDue")=="")?"0":Request.Form("AmountDue"));
	var AmountReceived = ((Request.Form("AmountDue")=="")?"0":Request.Form("AmountReceived"));	
	var PaidByEmployer = ((Request.Form("PaidByEmployer")=="1")?"1":"0");	
	var PaidByClient = ((Request.Form("PaidByClient")=="1")?"1":"0");	
	var PaidByEPPDConsultant = ((Request.Form("PaidByEPPDConsultant")=="1")?"1":"0");			
	var rsFollowUp = Server.CreateObject("ADODB.Recordset");
	rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
	rsFollowUp.Source="{call dbo.cp_follow_up("+Request.Form("MM_recordId")+",'2',"+ Request.QueryString("intAdult_id") +",0,'',0,0,0,'','',0,0,0,0,'',0,'"+Request.Form("BuyoutDueDate")+"','"+Request.Form("InvoiceNumber")+"','"+Request.Form("CaseNumber")+"',"+Request.Form("Case")+",'"+Request.Form("DateInvoiceSent")+"',"+AmountDue+","+AmountReceived+",'"+Request.Form("DateReceived")+"',"+PaidByEmployer+","+PaidByClient+","+PaidByEPPDConsultant+",'"+Request.Form("DefaultedDate")+"','"+Request.Form("DateCleared")+"','',0,'','',0,'E',0)}";
	rsFollowUp.CursorType = 0;
	rsFollowUp.CursorLocation = 2;
	rsFollowUp.LockType = 3;
	rsFollowUp.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m001q1102.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsFollowUpType = Server.CreateObject("ADODB.Recordset");
rsFollowUpType.ActiveConnection = MM_cnnASP02_STRING;
rsFollowUpType.Source = "{call dbo.cp_ASP_Lkup(94)}";
rsFollowUpType.CursorType = 0;
rsFollowUpType.CursorLocation = 2;
rsFollowUpType.LockType = 3;
rsFollowUpType.Open();

var rsFollowUp = Server.CreateObject("ADODB.Recordset");
rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
rsFollowUp.Source = "{call dbo.cp_Follow_up("+ Request.QueryString("intFlwup_id") + ",'2',0,0,'',0,0,0,'','',0,0,0,0,'',0, '','','',0,'',0.00,0.00,'',0,0,0,'','','','','','',1,'Q',0)}";
rsFollowUp.CursorType = 0;
rsFollowUp.CursorLocation = 2;
rsFollowUp.LockType = 3;
rsFollowUp.Open();
%>
<html>
<head>
	<title>Update EPPD Buyout Follow-Up</title>
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
				document.frm1102.reset();
			break;
		   	case 76 :
				//alert("L");
				history.back();
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
		if (!CheckDate(document.frm1102.DateInvoiceSent.value)) {
			alert("Invalid Date Invoice Sent.");
			document.frm1102.DateInvoiceSent.focus();
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
			alert("Invalid Date Received.");
			document.frm1102.DateReceived.focus();
			return ;
		}
		if (!CheckDate(document.frm1102.DefaultedDate.value)) {
			alert("Invalid Defaulted date.");
			document.frm1102.DefaultedDate.focus();
			return ;
		}
		if (!CheckDate(document.frm1102.DateCleared.value)) {
			alert("Invalid DateCleared.");
			document.frm1102.DateCleared.focus();
			return ;
		}
		document.frm1102.submit();
	}
	</script>	
</head>
<body onLoad="document.frm1102.BuyoutDueDate.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm1102">
<h5>Update EPPD Buyout Follow-Up</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Buyout Due Date:</td>
		<td nowrap>
			<input type="text" name="BuyoutDueDate" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsBOdue_date").Value)%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
		<td nowrap>Invoice Number:</td>
		<td nowrap><input type="text" name="InvoiceNumber" value="<%=(rsFollowUp.Fields.Item("chrInvoice_no").Value)%>" maxlength="10" tabindex="2"></td>
    </tr>
    <tr>
		<td nowrap>Date Invoice Sent:</td>
		<td nowrap>
			<input type="text" name="DateInvoiceSent" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsInvSend_date").Value)%>" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
		<td nowrap>Case:</td>
		<td nowrap>
			<select name="Case" tabindex="4">
				<option value="1" <%=((rsFollowUp.Fields.Item("bitIscase_new").Value == 1)?"SELECTED":"")%>>New
				<option value="0" <%=((rsFollowUp.Fields.Item("bitIscase_new").Value == 0)?"SELECTED":"")%>>Outstanding
			</select>
			Number:
			<input type="text" name="CaseNumber" value="<%=(rsFollowUp.Fields.Item("chvcase_no").Value)%>" maxlength="10" tabindex="5"></td>
    </tr>
    <tr> 
		<td nowrap>Amount Due:</td>
		<td nowrap>$<input type="text" name="AmountDue" value="<%=(rsFollowUp.Fields.Item("fltAmtDue").Value)%>" maxlength="8" size="8" tabindex="6" onKeypress="AllowNumericOnly();"></td>
	</tr>
    <tr> 
		<td nowrap>Amount Received:</td>
		<td nowrap>$<input type="text" name="AmountReceived" value="<%=(rsFollowUp.Fields.Item("fltPayAmtRx").Value)%>" maxlength="8" size="8" tabindex="7" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsPayRxdate").Value)%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Paid in Full by:</td>
		<td nowrap>
		  	<input type="checkbox" name="PaidByEmployer" <%=((rsFollowUp.Fields.Item("bitPayby_E").Value == 1)?"CHECKED":"")%> value="1" tabindex="9" class="chkstyle">Employer
			<input type="checkbox" name="PaidByClient" <%=((rsFollowUp.Fields.Item("bitPayby_C").Value == 1)?"CHECKED":"")%> value="1" tabindex="10" class="chkstyle">Client
			<input type="checkbox" name="PaidByEPPDConsultant" <%=((rsFollowUp.Fields.Item("bitPayby_V").Value == 1)?"CHECKED":"")%> value="1" tabindex="11" class="chkstyle">EPPD Consultant
		</td>
    </tr>
    <tr> 
		<td nowrap>Defaulted Date:</td>
		<td nowrap>
			<input type="text" name="DefaultedDate" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsDefault").Value)%>" size="11" maxlength="10" tabindex="12" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Date Cleared:</td>
		<td nowrap>
			<input type="text" name="DateCleared" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsCleared").Value)%>" size="11" maxlength="10" tabindex="13" accesskey="L" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="14" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="15" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="16" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsFollowUp.Fields.Item("intFlwup_id").Value%>">
</form>
</body>
</html>
<%
rsFollowUpType.Close();
rsFollowUp.Close();
%>