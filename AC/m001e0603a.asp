<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_update")) == "true"){
	var StartDate = ((String(Request.Form("StartDate"))=="undefined")?"1/1/1900":Request.Form("StartDate"));
	var EndDate = ((String(Request.Form("EndDate"))=="undefined")?"1/1/1900":Request.Form("EndDate"));

	var SFASComments = String(Request.Form("SFASComments")).replace(/'/g, "'");		
	var Comments = String(Request.Form("Comments")).replace(/'/g, "'");		
	var cmdInsertFinancialEligibility = Server.CreateObject("ADODB.Command");
	cmdInsertFinancialEligibility.ActiveConnection = MM_cnnASP02_STRING;
	cmdInsertFinancialEligibility.CommandText = "dbo.cp_List_Elgbty_Period2";
	cmdInsertFinancialEligibility.CommandType = 4;
	cmdInsertFinancialEligibility.CommandTimeout = 0;
	cmdInsertFinancialEligibility.Prepared = true;
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("RETURN_VALUE", 3, 4));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@intID", 3, 1,1,0));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@intAdult_id", 3, 1,1,Request.QueryString("intAdult_id")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@insFunding_Source_id", 2, 1,1,Request.Form("FundingSource")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@dtmEligibility_start", 200, 1,30,StartDate));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@dtmEligibility_end", 200, 1,30,EndDate));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@fltGrantAmt", 5, 1,1,Request.Form("TotalGrantAmount")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@fltGrantAmt_Tech", 5, 1,1,Request.Form("AmountUsedForTechnology")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@fltGrantAmt_Srv", 5, 1,1,Request.Form("AmountUsedForServices")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@insGrn_Qlf_Src_id", 2, 1,1,Request.Form("QualificationSource")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@chvNote", 200, 1,256,Comments));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@bitIs_Grnt_Eligible", 2, 1,1,Request.Form("GrantEligibility")));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@chvSFAS_comments", 200, 1,256,SFASComments));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@insMode", 2, 1,1,0));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@chrTask", 129, 1,1,'A'));
	cmdInsertFinancialEligibility.Parameters.Append(cmdInsertFinancialEligibility.CreateParameter("@insRtnValue", 3, 2));
	cmdInsertFinancialEligibility.Execute();

	var FinancialEligibilityID = cmdInsertFinancialEligibility.Parameters.Item("@insRtnValue").Value
	
	var rsGrantIneligibility = Server.CreateObject("ADODB.Recordset");
	rsGrantIneligibility.ActiveConnection = MM_cnnASP02_STRING;
	rsGrantIneligibility.CursorType = 0;
	rsGrantIneligibility.CursorLocation = 2;
	rsGrantIneligibility.LockType = 3;	
//  delete
//	rsGrantIneligibility.Source = "{call dbo.cp_grnt_inegbty(" + FinancialEligibilityID + ",0,'',0,'D',0)}";
//	rsGrantIneligibility.Open();
	//add
	for (var i=1; i <= Request.Form("Count"); i++){
		var DateResolved  = ((String(Request.Form("DateResolved")(i))=="undefined")?"1/1/1900":Request.Form("DateResolved")(i));
		rsGrantIneligibility.Source = "{call dbo.cp_grnt_inegbty(" + FinancialEligibilityID + "," + Request.Form("Reason")(i) + ",'" + DateResolved + "',0,'A',0)}";
		rsGrantIneligibility.Open();
//		rsGrantIneligibility.Close();
	}
	Response.Redirect("InsertSuccessful.html");
}

var rsFundingSource = Server.CreateObject("ADODB.Recordset");
rsFundingSource.ActiveConnection = MM_cnnASP02_STRING;
rsFundingSource.Source = "{call dbo.cp_funding_source_attributes(0,0,1,0,1,0,0,0,2,'Q',0)}";
rsFundingSource.CursorType = 0;
rsFundingSource.CursorLocation = 2;
rsFundingSource.LockType = 3;
rsFundingSource.Open();

var rsQualificationSource = Server.CreateObject("ADODB.Recordset");
rsQualificationSource.ActiveConnection = MM_cnnASP02_STRING;
rsQualificationSource.Source = "{call dbo.cp_grant_qlf_src(0,'',0,'Q',0)}";
rsQualificationSource.CursorType = 0;
rsQualificationSource.CursorLocation = 2;
rsQualificationSource.LockType = 3;
rsQualificationSource.Open();

var rsReason = Server.CreateObject("ADODB.Recordset");
rsReason.ActiveConnection = MM_cnnASP02_STRING;
rsReason.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,6,'',2,'Q',0)}";
rsReason.CursorType = 0;
rsReason.CursorLocation = 2;
rsReason.LockType = 3;
rsReason.Open();%>
<html>
<head>
	<title>New Financial Eligibility Information</title>
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
		   	case 76 :
				//alert("L");
				window.close();
			break;
		}
	}
	</script>	
	<script language="Javascript">
	var count = 0;	
	function Save(){
		if (isNaN(document.frm0603.TotalGrantAmount.value)){
			alert("Invalid Total Grant Amount.");
			document.frm0603.TotalGrantAmount.focus();
			return ;
		}
		if (isNaN(document.frm0603.AmountUsedForTechnology.value)){
			alert("Invalid Amount Used For Technology.");
			document.frm0603.AmountUsedForTechnology.focus();
			return ;
		}
		
		if (isNaN(document.frm0603.AmountUsedForServices.value)){
			alert("Invalid Amount Used For Services.");
			document.frm0603.AmountUsedForServices.focus();
			return ;
		}
		
		if (!CheckDate(document.frm0603.StartDate.value)){
			alert("Invalid Start Date.");
			document.frm0603.StartDate.focus();
			return ;
		}
		
		if (!CheckDate(document.frm0603.EndDate.value)){
			alert("Invalid End Date.");
			document.frm0603.EndDate.focus();
			return ;		
		}

		if (!CheckDateBetween(Trim(document.frm0603.StartDate.value)+" and "+Trim(document.frm0603.EndDate.value))) {
			alert("Start Date is after End Date.");
			document.frm0603.EndDate.focus();
			return ;
		}

		if (document.frm0603.SFASComments.value.length > 256) {
			alert("SFAS Comments cannot exceed 256 characters");
			document.frm0603.SFASComments.focus();
			return ;
		}		

		if (document.frm0603.Comments.value.length > 256) {
			alert("Comments cannot exceed 256 characters");
			document.frm0603.Comments.focus();
			return ;
		}		

		if (document.frm0603.GrantEligibility.value=="0"){
			for (var i = 0; i < count; i++) {
				if (!CheckDate(document.frm0603.DateResolved[i].value)) {
					alert("Invalid Date Resolved");
					document.frm0603.DateResolved[i].focus();
					return; 
				}
			}
		}
		
		if (document.frm0603.TotalGrantAmount.value=="") document.frm0603.TotalGrantAmount.value = 0;
		if (document.frm0603.AmountUsedForTechnology.value=="") document.frm0603.AmountUsedForTechnology.value = 0;
		if (document.frm0603.AmountUsedForServices.value=="") document.frm0603.AmountUsedForServices.value = 0;
		
		document.frm0603.Count.value=count;
		document.frm0603.submit();
	}
	
	function ChangeAmount(){
		document.frm0603.GrantAmountLeft.value=FormatCurrency(document.frm0603.TotalGrantAmount.value - document.frm0603.AmountUsedForServices.value - document.frm0603.AmountUsedForTechnology.value);
	}
	
	function ChangeEligibility(){
		if (document.frm0603.GrantEligibility.value=="1"){
			for (var i = 0; i < 4; i++){
				document.frm0603.Reason[i].style.visibility = "hidden";
				document.frm0603.DateResolved[i].style.visibility = "hidden";			
				document.frm0603.DateResolved[i].value = "";						
			}
			document.frm0603.Add.disabled = true;
			document.frm0603.Remove.disabled = true;
			count = 0;
			document.frm0603.StartDate.disabled = false;
			document.frm0603.EndDate.disabled = false;
		} else {
			if (count == 0) {
				document.frm0603.Remove.disabled = true;
				document.frm0603.Add.disabled = false;				
			}
			if (count == 4) {
				document.frm0603.Remove.disabled = false;			
				document.frm0603.Add.disabled = true;
			}
			document.frm0603.StartDate.disabled = true;
			document.frm0603.EndDate.disabled = true;
			document.frm0603.StartDate.value = "";
			document.frm0603.EndDate.value = "";			
		}
	}	

	function AddReason(){
		document.frm0603.Reason[count].style.visibility = "visible";
		document.frm0603.DateResolved[count].style.visibility = "visible";
		count++;
		document.frm0603.Remove.disabled = false;
		if (count == 4) document.frm0603.Add.disabled = true;
	}
	
	function RemoveReason(){
		count--;
		document.frm0603.Reason[count].style.visibility = "hidden";
		document.frm0603.DateResolved[count].style.visibility = "hidden";
		document.frm0603.DateResolved[count].value = "";		
		document.frm0603.Add.disabled = false;
		if (count == 0) document.frm0603.Remove.disabled = true;
	}
	
	function Init(){
		document.frm0603.Add.disabled = true;
		document.frm0603.Remove.disabled = true;	
		document.frm0603.QualificationSource.focus()
	}	
	</script>
</head>
<body onLoad="Init();">
<form name="frm0603" method="POST" action="<%=MM_editAction%>">
<h5>New Financial Eligibility Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Qualification Source:</td>
		<td nowrap><select name="QualificationSource" tabindex="1" accesskey="F" style="width:180px">
		<%
		while (!rsQualificationSource.EOF) {
		%>
			<option value="<%=rsQualificationSource.Fields.Item("insGrant_Qlf_Src").Value%>"><%=rsQualificationSource.Fields.Item("chvGrant_Qlf_Src").Value%>
		<%
			rsQualificationSource.MoveNext();
		}
		%>
		</select></td>
    </tr>
	<tr>
		<td nowrap>Funding Source:</td>
		<td nowrap><select name="FundingSource" tabindex="2" style="width:180px">
		<%
		while (!rsFundingSource.EOF) {
		%>
			<option value="<%=rsFundingSource.Fields.Item("insFunding_source_id").Value%>"><%=rsFundingSource.Fields.Item("chvfunding_source_name").Value%>
		<%
			rsFundingSource.MoveNext();
		}
		%>
		</select></td>
    </tr>
	<tr>
		<td nowrap>Grant Eligibility:</td>
		<td nowrap><select name="GrantEligibility" tabindex="3" onChange="ChangeEligibility();" style="width:180px">
			<option value="1">Eligible
			<option value="0">Not Eligible
		</select></td>
	</tr>
	<tr>
		<td nowrap valign="top">Ineligibile Reasons:</td>
		<td valign="top">
			<table cellpadding="1" cellspacing="1" style="border: 1px solid">
				<tr>
					<td class="headrow" align="center">Reason</td>
					<td class="headrow" align="center">Date Resolved</td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="4" style="width: 200px; visibility='hidden'">
						<%
						while (!rsReason.EOF) {
						%>
							<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>"><%=rsReason.Fields.Item("chvDocDesc").Value%>
						<%
							rsReason.MoveNext();
						}
						rsReason.MoveFirst();
						%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" tabindex="5" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="6" style="width: 200px; visibility='hidden'">
						<%
						while (!rsReason.EOF) {
						%>
							<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>"><%=rsReason.Fields.Item("chvDocDesc").Value%>
						<%
							rsReason.MoveNext();
						}
						rsReason.MoveFirst();
						%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" tabindex="7" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="8" style="width: 200px; visibility='hidden'">
						<%
						while (!rsReason.EOF) {
						%>
							<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>"><%=rsReason.Fields.Item("chvDocDesc").Value%>
						<%
							rsReason.MoveNext();
						}
						rsReason.MoveFirst();
						%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" tabindex="9" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="10" style="width: 200px; visibility='hidden'">
						<%
						while (!rsReason.EOF) {
						%>
							<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>"><%=rsReason.Fields.Item("chvDocDesc").Value%>
						<%
							rsReason.MoveNext();
						}
						rsReason.MoveFirst();
						%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" tabindex="11" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
			</table>
		</td>	
	</tr>
	<tr>
		<td></td>
		<td nowrap>
			<input type="button" name="Add" value="Add Reason" onClick="AddReason();" tabindex="12" class="btnstyle">&nbsp;
			<input type="button" name="Remove" value="Remove Reason" onClick="RemoveReason();" tabindex="13" class="btnstyle">
		</td>		
	</tr>	
	<tr>
		<td nowrap valign="top">SFAS Comments:</td>
		<td nowrap><textarea name="SFASComments" tabindex="14" cols="60" rows="4"></textarea></td>
	</tr>	
	<tr> 
		<td nowrap>Start Date:</td>
		<td nowrap>
			<input type="text" name="StartDate" size="11" maxlength="10" tabindex="15" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap>End Date:</td>
		<td nowrap>
			<input type="text" name="EndDate" size="11" maxlength="10" tabindex="16" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Total Grant Amount:</td>
		<td nowrap>$<input type="text" name="TotalGrantAmount" value="8000.00" size="9" onChange="ChangeAmount();" tabindex="17" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td nowrap align="right"></td>
		<td nowrap>
			$<input type="text" name="AmountUsedForTechnology" onChange="ChangeAmount();" size="9" tabindex="18" onKeypress="AllowNumericOnly();">
			<span style="font-size: 7pt">(Estimate Amount Used For Technology)</span>
		</td>
	</tr>
	<tr>
		<td nowrap align="right"></td>
		<td nowrap>
			$<input type="text" name="AmountUsedForServices" onChange="ChangeAmount();" size="9" tabindex="19" onKeypress="AllowNumericOnly();">
			<span style="font-size: 7pt">(Estimate Amount Used For Services)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Grant Amount Left:</td>
		<td nowrap><input type="text" name="GrantAmountLeft" value="$0.00" size="12" tabindex="20" onKeypress="AllowNumericOnly();" readonly></td>
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap><textarea name="Comments" tabindex="21" accesskey="L" cols="60" rows="4"></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="22" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="23" onClick="window.close()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="Count" value="0">
</form>
</body>
</html>
<%
rsReason.Close();
rsFundingSource.Close();
rsQualificationSource.Close();
%>