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
	var CIPRequired = ((Request.Form("CIPRequired")=="1")?"1":"0");	
	var ReturnEquipment = ((Request.Form("ReturnEquipment")=="1")?"1":"0");		
	var ExpectedEducationalProgramCompletionDate = ((String(Request.Form("ExpectedEducationalProgramCompletionDate"))=="undefined")?"":Request.Form("ExpectedEducationalProgramCompletionDate"));
	var VocationalLoanEndDate = ((String(Request.Form("VocationalLoanEndDate"))=="undefined")?"":Request.Form("VocationalLoanEndDate"));
	var BuyOutInPlace = ((Request.Form("BuyOutInPlace")=="1")?"1":"0");
	var ReturnBy = ((String(Request.Form("ReturnBy"))=="undefined")?"":Request.Form("ReturnBy"));
	var IssueState = ((Request.Form("IssueState")=="1")?"1":"0");	
	var rsFollowUp = Server.CreateObject("ADODB.Recordset");
	rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
	rsFollowUp.Source = "{call dbo.cp_follow_up("+Request.Form("MM_recordId")+",'1',"+ Request.QueryString("intAdult_id") +","+Request.Form("Year")+",'"+Request.Form("DateReceived")+"',"+Request.Form("Type")+","+Request.Form("EquipmentNeedsMet")+","+Request.Form("DisabilityChanged")+",'"+ExpectedEducationalProgramCompletionDate+"','"+VocationalLoanEndDate+"',"+BuyOutInPlace+","+Request.Form("ActionRequired")+","+CIPRequired+","+ReturnEquipment+",'"+ReturnBy+"',"+IssueState+",'','','',0,'',0,0,'',0,0,0,'','','',0,'','',0,'E',0)}";
	rsFollowUp.CursorType = 0;
	rsFollowUp.CursorLocation = 2;
	rsFollowUp.LockType = 3;
	rsFollowUp.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m001q1101.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsFollowUp = Server.CreateObject("ADODB.Recordset");
rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
rsFollowUp.Source = "{call dbo.cp_Follow_up("+ Request.QueryString("intFlwup_id") + ",'1',0,0,'',0,0,0,'','',0,0,0,0,'',0, '','','',0,'',0.00,0.00,'',0,0,0,'','','','','','',1,'Q',0)}";
rsFollowUp.CursorType = 0;
rsFollowUp.CursorLocation = 2;
rsFollowUp.LockType = 3;
rsFollowUp.Open();
%>
<html>
<head>
	<title>Update Follow-Up</title>
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
			case 85:
				//alert("U");
				document.frm1101.reset();
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
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm1101">
<h5>Update Annual Follow-Up</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="1" onChange="ChangeType();" accesskey="F">
			<option value="1" <%=((rsFollowUp.Fields.Item("bitIs_Educatinal").Value == 1)?"SELECTED":"")%>>Educational
			<option value="0" <%=((rsFollowUp.Fields.Item("bitIs_Educatinal").Value == 0)?"SELECTED":"")%>>Vocational
		</select></td>
    </tr>
    <tr> 
		<td nowrap>Year:</td>
		<td nowrap><input type="text" name="Year" value="<%=(rsFollowUp.Fields.Item("insYear").Value)%>" maxlength="4" size="4" onKeypress="AllowNumericOnly();" tabindex="2" ></td>
	</tr>
	<tr>
		<td nowrap>Date Received:</td>
		<td nowrap>
			<input type="text" name="DateReceived" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsRx_date").Value)%>" maxlength="10" size="11" tabindex="3" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Services/equipment<br>meeting needs:</td>
		<td nowrap><select name="EquipmentNeedsMet" tabindex="4">
			<option value="1" <%=((rsFollowUp.Fields.Item("bitSrvEqu_need").Value == 1)?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsFollowUp.Fields.Item("bitSrvEqu_need").Value == 0)?"SELECTED":"")%>>No
		</select></td>	  
	</tr>
    <tr> 
		<td nowrap>Disability changed:</td>
		<td nowrap><select name="DisabilityChanged" tabindex="5">
			<option value="1" <%=((rsFollowUp.Fields.Item("bitDsbChanged").Value == 1)?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsFollowUp.Fields.Item("bitDsbChanged").Value == 0)?"SELECTED":"")%>>No
		</select></td>	  
    <tr> 
		<td nowrap><div id="EEPCDLabel">Expected educational<br>program completion date:</div></td>
		<td nowrap>
			<input type="text" name="ExpectedEducationalProgramCompletionDate" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsEdCmplDate").Value)%>" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)" >
			<span id="EEPCDLabel2" style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap><div id="VLEDLabel">Vocational loan end date:</div></td>
		<td nowrap>
			<input type="text" name="VocationalLoanEndDate" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsVocLoanDate").Value)%>" size="11" maxlength="10" tabindex="7" onChange="FormatDate(this)" >
			<span id="VLEDLabel2" style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap><div id="BIPLabel">Buyout in place:</div></td>
		<td nowrap><select name="BuyoutInPlace" tabindex="8">
			<option value="1" <%=((rsFollowUp.Fields.Item("bitBuyOut").Value == 1)?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsFollowUp.Fields.Item("bitBuyOut").Value == 0)?"SELECTED":"")%>>No
		</select></td>	  	  
    </tr>
    <tr> 
		<td nowrap>Action required:</td>
		<td nowrap><select name="ActionRequired" onChange="ChangeActionRequired();" tabindex="9">
			<option value="1" <%=((rsFollowUp.Fields.Item("bitAction_Req").Value == 1)?"SELECTED":"")%>>Yes
			<option value="0" <%=((rsFollowUp.Fields.Item("bitAction_Req").Value == 0)?"SELECTED":"")%>>No
		</select></td>	  
    </tr>
    <tr> 
		<td></td>
		<td nowrap><div id="ActionRequiredBlock">		
			<input type="checkbox" name="CIPRequired" <%=((rsFollowUp.Fields.Item("bitCIP_Req").Value == 1)?"CHECKED":"")%> value="1" tabindex="10" class="chkstyle">CIP required<br>
			<input type="checkbox" name="ReturnEquipment" <%=((rsFollowUp.Fields.Item("bitRtn_Eqp").Value == 1)?"CHECKED":"")%> value="1" tabindex="11" class="chkstyle">Return Equipment<br>
			By:&nbsp;<input type="text" name="ReturnBy" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsAction_date").Value)%>" size="11" maxlength="10" tabindex="12" onChange="FormatDate(this)"><span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</div></td>
	</tr>
    <tr> 
		<td nowrap><div id="IssueStateBlock">Issue State:</div></td>
		<td nowrap><select name="IssueState" tabindex="13" accesskey="L">
			<option value="1" <%=((rsFollowUp.Fields.Item("bitissue").Value == 1)?"SELECTED":"")%>>Resolved
			<option value="0" <%=((rsFollowUp.Fields.Item("bitissue").Value == 0)?"SELECTED":"")%>>Unresolved
		</select></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="14" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="15" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="16" onClick="history.back();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsFollowUp.Fields.Item("intFlwup_id").Value%>">
</form>
</body>
</html>
<%
rsFollowUp.Close();
%>
