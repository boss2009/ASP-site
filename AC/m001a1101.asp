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
	var CIPRequired = ((Request.Form("CIPRequired")=="1")?"1":"0");	
	var ReturnEquipment = ((Request.Form("ReturnEquipment")=="1")?"1":"0");		
	var ExpectedEducationalProgramCompletionDate = ((String(Request.Form("ExpectedEducationalProgramCompletionDate"))=="undefined")?"":Request.Form("ExpectedEducationalProgramCompletionDate"));
	var VocationalLoanEndDate = ((String(Request.Form("VocationalLoanEndDate"))=="undefined")?"":Request.Form("VocationalLoanEndDate"));
	var BuyoutInPlace = ((Request.Form("BuyoutInPlace")=="1")?"1":"0");
	var ReturnBy = ((String(Request.Form("ReturnBy"))=="undefined")?"":Request.Form("ReturnBy"));
	var IssueState = ((Request.Form("IssueState")=="1")?"1":"0");	
	var rsFollowUp = Server.CreateObject("ADODB.Recordset");
	rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
	rsFollowUp.Source="{call dbo.cp_follow_up(0,'1',"+ Request.QueryString("intAdult_id") +","+Request.Form("Year")+",'"+Request.Form("DateReceived")+"',"+Request.Form("Type")+","+Request.Form("ServiceMeetingNeeds")+","+Request.Form("DisabilityChanged")+",'"+ExpectedEducationalProgramCompletionDate+"','"+VocationalLoanEndDate+"',"+BuyoutInPlace+","+Request.Form("ActionRequired")+","+CIPRequired+","+ReturnEquipment+",'"+ReturnBy+"',"+IssueState+",'','','',0,'',0,0,'',0,0,0,'','','',0,'','',0,'A',0)}";
	rsFollowUp.CursorType = 0;
	rsFollowUp.CursorLocation = 2;
	rsFollowUp.LockType = 3;
	rsFollowUp.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New Annual Follow-Up</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js""></script>
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
		if (!CheckDate(document.frm1101.DateReceived.value)) {
			alert("Invalid Date Received.");
			document.frm1101.DateReceived.focus();
			return ;
		}
		if (!CheckDate(document.frm1101.ExpectedEducationalProgramCompletionDate.value)) {
			alert("Invalid Expected Completion Date.");
			document.frm1101.ExpectedEducationalProgramCompletionDate.focus();
			return ;
		}
		if (!CheckDate(document.frm1101.VocationalLoanEndDate.value)) {
			alert("Invalid Vocational Loan End Date.");
			document.frm1101.VocationalLoanEndDate.focus();
			return ;
		}
		if (!CheckDate(document.frm1101.ReturnBy.value)) {
			alert("Invalid Return By Date.");
			document.frm1101.ReturnBy.focus();
			return ;
		}
		document.frm1101.submit();
	}
	
	function ChangeType(){
		if (document.frm1101.Type.value==1){
			document.frm1101.ExpectedEducationalProgramCompletionDate.disabled = false;					
			EEPCDLabel.style.visibility = "visible";			
			EEPCDLabel2.style.visibility = "visible";						
			document.frm1101.ExpectedEducationalProgramCompletionDate.style.visibility = "visible";								
			document.frm1101.VocationalLoanEndDate.disabled = true;
			VLEDLabel.style.visibility = "hidden";
			VLEDLabel2.style.visibility = "hidden";			
			document.frm1101.VocationalLoanEndDate.style.visibility = "hidden";						
			document.frm1101.BuyoutInPlace.disabled = true;
			BIPLabel.style.visibility = "hidden";
			document.frm1101.BuyoutInPlace.style.visibility = "hidden";
		} else {			
			document.frm1101.ExpectedEducationalProgramCompletionDate.disabled = true;
			EEPCDLabel.style.visibility = "hidden";
			EEPCDLabel2.style.visibility = "hidden";			
			document.frm1101.ExpectedEducationalProgramCompletionDate.style.visibility = "hidden";
			document.frm1101.VocationalLoanEndDate.disabled = false;			
			VLEDLabel.style.visibility = "visible";
			VLEDLabel2.style.visibility = "visible";			
			document.frm1101.VocationalLoanEndDate.style.visibility = "visible";
			document.frm1101.BuyoutInPlace.disabled = false;			
			BIPLabel.style.visibility = "visible";
			document.frm1101.BuyoutInPlace.style.visibility = "visible";
		}		
	}
	
	function ChangeActionRequired(){
		if (document.frm1101.ActionRequired.value==1){
			document.frm1101.ReturnBy.disabled = false;
			document.frm1101.CIPRequired.disabled = false;
			document.frm1101.ReturnEquipment.disabled = false;
			document.frm1101.IssueState.disabled = false;
			ActionRequiredBlock.style.visibility = "visible";
			IssueStateBlock.style.visibility = "visible";
			document.frm1101.IssueState.style.visibility = "visible";			
		} else {
			document.frm1101.ReturnBy.disabled = true;
			document.frm1101.CIPRequired.disabled = true;			
			document.frm1101.ReturnEquipment.disabled = true;			
			document.frm1101.IssueState.disabled = true;
			ActionRequiredBlock.style.visibility = "hidden";			
			IssueStateBlock.style.visibility = "hidden";			
			document.frm1101.IssueState.style.visibility = "hidden";						
		}
	}
	
	function Init(){
		ChangeType();
		ChangeActionRequired();
		document.frm1101.Type.focus();
	}	
	</script>
</head>
<body onLoad="Init();">
<form name="frm1101" method="POST" action="<%=MM_editAction%>">
<h5>New Annual Follow-Up</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="1" onChange="ChangeType();" accesskey="F">
			<option value="1" SELECTED>Educational
			<option value="0">Vocational
		</select></td> 
    </tr>
    <tr>
		<td nowrap>Year:</td>
		<td nowrap><input type="text" name="Year" value="<%=CurrentYear()%>" onKeypress="AllowNumericOnly();" size="4" maxlength="4" tabindex="2"></td>
	</tr>
	<tr>
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap colspan="2">
			Services/equipment meeting needs:
			<select name="ServiceMeetingNeeds" tabindex="4">
				<option value="1" SELECTED>Yes
				<option value="0">No
			</select>
		</td>
    </tr>
    <tr> 
		<td nowrap colspan="2">
			Disability changed:
			<select name="DisabilityChanged" tabindex="5">
				<option value="1">Yes
				<option value="0" SELECTED>No
			</select>
		</td> 
    </tr>
    <tr> 
		<td nowrap colspan="2">
			<div id="EEPCDLabel">Expected educational program completion date:
			<input type="text" name="ExpectedEducationalProgramCompletionDate" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)">
			<span id="EEPCDLabel2" style="font-size: 7pt">(mm/dd/yyyy)</span>
			</div>
		</td>
    </tr>
    <tr>
		<td nowrap colspan="2">
			<div id="VLEDLabel">Vocational loan end date:
			<input type="text" name="VocationalLoanEndDate" size="11" maxlength="10" tabindex="7" onChange="FormatDate(this)">
			<span id="VLEDLabel2" style="font-size: 7pt">(mm/dd/yyyy)</span>
			</div>			
		</td>
    </tr>
    <tr>
		<td nowrap colspan="2">
			<div id="BIPLabel">Buyout in place:
			<select name="BuyoutInPlace" tabindex="8">
				<option value="1" SELECTED>Yes
				<option value="0">No
			</select>
			</div>
		</td>			
	</tr>
    <tr> 
		<td nowrap>Action required:</td>
		<td nowrap><select name="ActionRequired" tabindex="9" onChange="ChangeActionRequired();">
			<option value="1" SELECTED>Yes
			<option value="0">No
		</select></td>
	</tr>
    <tr> 
		<td></td>
		<td nowrap><div id="ActionRequiredBlock">
			<input type="checkbox" name="CIPRequired" value="1" tabindex="10" class="chkstyle">CIP required<br>
			<input type="checkbox" name="ReturnEquipment" value="1" tabindex="11" class="chkstyle">Return Equipment<br>
			By&nbsp;<input type="text" name="ReturnBy" size="11" maxlength="10" tabindex="12" onChange="FormatDate(this)"><span style="font-size: 7pt">(mm/dd/yyyy)</span><br>
		</div></td>
	</tr>
    <tr> 	
		<td nowrap><div id="IssueStateBlock">Issue State:</div></td>
		<td nowrap><select name="IssueState" tabindex="13" accesskey="L">
			<option value="1">Resolved
			<option value="0" SELECTED>Unresolved
		</select></td> 
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