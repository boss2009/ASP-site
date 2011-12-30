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
	var SFASComments = String(Request.Form("SFASComments")).replace(/'/g, "'");		
	var Comments = String(Request.Form("Comments")).replace(/'/g, "'");		
	var StartDate = ((String(Request.Form("StartDate"))=="undefined")?"1/1/1900":Request.Form("StartDate"));
	var EndDate = ((String(Request.Form("EndDate"))=="undefined")?"1/1/1900":Request.Form("EndDate"));
	
	var rsFinancialEligibility = Server.CreateObject("ADODB.Recordset");
	rsFinancialEligibility.ActiveConnection = MM_cnnASP02_STRING;
	rsFinancialEligibility.Source = "{call dbo.cp_List_Elgbty_Period2("+Request.QueryString("insGrnt_Elgbty")+","+Request.QueryString("intAdult_id")+","+Request.Form("FundingSource")+",'"+StartDate+"','"+EndDate+"',"+Request.Form("TotalGrantAmount")+","+Request.Form("AmountUsedForTechnology")+","+Request.Form("AmountUsedForServices")+","+Request.Form("QualificationSource")+",'"+Comments+"',"+Request.Form("GrantEligibility")+",'"+SFASComments+"',0,'E',0)}";
	rsFinancialEligibility.CursorType = 0;
	rsFinancialEligibility.CursorLocation = 2;
	rsFinancialEligibility.LockType = 3;
	rsFinancialEligibility.Open();

	var rsGrantIneligibility = Server.CreateObject("ADODB.Recordset");
	rsGrantIneligibility.ActiveConnection = MM_cnnASP02_STRING;
	rsGrantIneligibility.CursorType = 0;
	rsGrantIneligibility.CursorLocation = 2;
	rsGrantIneligibility.LockType = 3;	
	//delete
	rsGrantIneligibility.Source = "{call dbo.cp_grnt_inegbty(" + Request.QueryString("insGrnt_Elgbty") + ",0,'',0,'D',0)}";
	rsGrantIneligibility.Open();
	//add
	for (var i=1; i<=Request.Form("Count"); i++){
		var DateResolved = ((String(Request.Form("DateResolved")(i))=="undefined")?"1/1/1900":Request.Form("DateResolved")(i));	
		rsGrantIneligibility.Source = "{call dbo.cp_grnt_inegbty(" + Request.QueryString("insGrnt_Elgbty") + "," + Request.Form("Reason")(i) + ",'" + DateResolved + "',0,'A',0)}";
		rsGrantIneligibility.Open();
	}
	Response.Redirect("m001e0603.asp?intReferral_id="+Request.QueryString("intReferral_id")+"&insGrnt_Elgbty="+Request.QueryString("insGrnt_Elgbty")+"&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsFinancialEligibility = Server.CreateObject("ADODB.Recordset");
rsFinancialEligibility.ActiveConnection = MM_cnnASP02_STRING;
rsFinancialEligibility.Source = "{call dbo.cp_List_Elgbty_Period2("+Request.QueryString("insGrnt_Elgbty")+",0,0,'','',0,0,0,0,'',0,'',1,'Q',0)}";
rsFinancialEligibility.CursorType = 0;
rsFinancialEligibility.CursorLocation = 2;
rsFinancialEligibility.LockType = 3;
rsFinancialEligibility.Open();

var rsGrantIneligibility = Server.CreateObject("ADODB.Recordset");
rsGrantIneligibility.ActiveConnection = MM_cnnASP02_STRING;
rsGrantIneligibility.Source = "{call dbo.cp_Grnt_Inegbty("+Request.QueryString("insGrnt_Elgbty")+",0,'',2,'Q',0)}";
rsGrantIneligibility.CursorType = 0;
rsGrantIneligibility.CursorLocation = 2;
rsGrantIneligibility.LockType = 3;
rsGrantIneligibility.Open();

var count = 0;
while (!rsGrantIneligibility.EOF) {
	count++;
	rsGrantIneligibility.MoveNext();
}
if (count > 0) rsGrantIneligibility.MoveFirst();

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
	<title>Update Financial Eligibility Information</title>
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
	var count = <%=count%>;	
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
		
		if (document.frm0603.GrantEligibility.value=="0"){
			for (var i = 0; i < count; i++) {
				if (!CheckDate(document.frm0603.DateResolved[i].value)) {
					alert("Invalid Date Resolved");
					document.frm0603.DateResolved[i].focus();
					return; 
				}
			}
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
		
		if (document.frm0603.TotalGrantAmount.value=="") document.frm0603.TotalGrantAmount.value = 0;
		if (document.frm0603.AmountUsedForTechnology.value=="") document.frm0603.AmountUsedForTechnology.value = 0;
		if (document.frm0603.AmountUsedForServices.value=="") document.frm0603.AmountUsedForServices.value = 0;
		
		document.frm0603.Count.value=count;
		document.frm0603.submit();
	}
	
	function ChangeAmount(){
		document.frm0603.GrantAmountLeft.value=FormatCurrency(document.frm0603.TotalGrantAmount.value - document.frm0603.AmountUsedForTechnology.value - document.frm0603.AmountUsedForServices.value);	
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
		ChangeEligibility();
		if (count > 0) {
			for (var i = 0; i < count; i++) {
				document.frm0603.Reason[i].style.visibility = "visible";
				document.frm0603.DateResolved[i].style.visibility = "visible";				
			}
		}
		document.frm0603.QualificationSource.focus()		
		ChangeAmount();
	}	
	</script>
</head>
<body onLoad="Init();">
<form name="frm0603" method="POST" action="<%=MM_editAction%>">
<h5>Update Financial Eligibility Information</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Qualification Source:</td>
		<td nowrap><select name="QualificationSource" tabindex="1" accesskey="F" style="width:180px">
		<%
		while (!rsQualificationSource.EOF) {
		%>
			<option value="<%=rsQualificationSource.Fields.Item("insGrant_Qlf_Src").Value%>" <%=((rsQualificationSource.Fields.Item("insGrant_Qlf_Src").Value==rsFinancialEligibility.Fields.Item("insGrn_Qlf_Src_id").Value)?"SELECTED":"")%>><%=rsQualificationSource.Fields.Item("chvGrant_Qlf_Src").Value%>
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
			<option value="<%=rsFundingSource.Fields.Item("insFunding_source_id").Value%>" <%=((rsFundingSource.Fields.Item("insFunding_source_id").Value==rsFinancialEligibility.Fields.Item("insFunding_Source_id").Value)?"SELECTED":"")%>><%=rsFundingSource.Fields.Item("chvfunding_source_name").Value%>
		<%
			rsFundingSource.MoveNext();
		}
		%>
		</select></td>
    </tr>	
	<tr>
		<td nowrap>Grant Eligibility:</td>
		<td nowrap><select name="GrantEligibility" tabindex="3" onChange="ChangeEligibility();" style="width:180px">
			<option value="1" <%=((rsFinancialEligibility.Fields.Item("bitIs_Grnt_Eligible").Value=="1")?"SELECTED":"")%>>Eligible
			<option value="0" <%=((rsFinancialEligibility.Fields.Item("bitIs_Grnt_Eligible").Value=="0")?"SELECTED":"")%>>Not Eligible
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
						<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>" <%if (!rsGrantIneligibility.EOF) { Response.Write(((rsReason.Fields.Item("intDoc_id").Value==rsGrantIneligibility.Fields.Item("intDoc_id").Value)?"SELECTED":""))}%>><%=rsReason.Fields.Item("chvDocDesc").Value%>
					<%
						rsReason.MoveNext();
					}
					rsReason.MoveFirst();
					%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" value="<%if (!rsGrantIneligibility.EOF) { Response.Write(FilterDate(rsGrantIneligibility.Fields.Item("dtsResolved_Date").Value)); rsGrantIneligibility.MoveNext();}%>" tabindex="5" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="6" style="width: 200px; visibility='hidden'">
					<%
					while (!rsReason.EOF) {
					%>
						<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>" <%if (!rsGrantIneligibility.EOF) { Response.Write(((rsReason.Fields.Item("intDoc_id").Value==rsGrantIneligibility.Fields.Item("intDoc_id").Value)?"SELECTED":""))}%>><%=rsReason.Fields.Item("chvDocDesc").Value%>
					<%
						rsReason.MoveNext();
					}
					rsReason.MoveFirst();
					%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" value="<%if (!rsGrantIneligibility.EOF) { Response.Write(FilterDate(rsGrantIneligibility.Fields.Item("dtsResolved_Date").Value)); rsGrantIneligibility.MoveNext();}%>" tabindex="5" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="8" style="width: 200px; visibility='hidden'">
					<%
					while (!rsReason.EOF) {
					%>
						<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>" <%if (!rsGrantIneligibility.EOF) { Response.Write(((rsReason.Fields.Item("intDoc_id").Value==rsGrantIneligibility.Fields.Item("intDoc_id").Value)?"SELECTED":""))}%>><%=rsReason.Fields.Item("chvDocDesc").Value%>
					<%
						rsReason.MoveNext();
					}
					rsReason.MoveFirst();
					%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" value="<%if (!rsGrantIneligibility.EOF) { Response.Write(FilterDate(rsGrantIneligibility.Fields.Item("dtsResolved_Date").Value)); rsGrantIneligibility.MoveNext();}%>" tabindex="5" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
				</tr>
				<tr>
					<td nowrap><select name="Reason" tabindex="10" style="width: 200px; visibility='hidden'">
					<%
					while (!rsReason.EOF) {
					%>
						<option value="<%=rsReason.Fields.Item("intDoc_id").Value%>" <%if (!rsGrantIneligibility.EOF) { Response.Write(((rsReason.Fields.Item("intDoc_id").Value==rsGrantIneligibility.Fields.Item("intDoc_id").Value)?"SELECTED":""))}%>><%=rsReason.Fields.Item("chvDocDesc").Value%>
					<%
						rsReason.MoveNext();
					}
					rsReason.MoveFirst();
					%>
					</select></td>
					<td nowrap align="center"><input type="text" name="DateResolved" value="<%if (!rsGrantIneligibility.EOF) { Response.Write(FilterDate(rsGrantIneligibility.Fields.Item("dtsResolved_Date").Value)); rsGrantIneligibility.MoveNext();}%>" tabindex="5" size="11" maxlength="10" style="visibility:'hidden'" onChange="FormatDate(this)"></td>
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
		<td nowrap valign="top"><textarea name="SFASComments" tabindex="14" cols="60" rows="4"><%=rsFinancialEligibility.Fields.Item("chvSFAS_comment").Value%></textarea></td>
	</tr>	
	<tr> 
		<td nowrap>Start Date:</td>
		<td nowrap>
			<input type="text" name="StartDate" value="<%=FilterDate(rsFinancialEligibility.Fields.Item("dtmEligibility_start").Value)%>" size="11" maxlength="10" tabindex="15" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr> 
		<td nowrap>End Date:</td>
		<td nowrap>
			<input type="text" name="EndDate" value="<%=FilterDate(rsFinancialEligibility.Fields.Item("dtmEligibility_end").Value)%>" size="11" maxlength="10" tabindex="16" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Total Grant Amount:</td>
		<td nowrap>$<input type="text" name="TotalGrantAmount" value="<%=rsFinancialEligibility.Fields.Item("fltGrantAmt").Value%>" size="9" onChange="ChangeAmount();" tabindex="17" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td nowrap></td>
		<td nowrap>
			$<input type="text" name="AmountUsedForTechnology" value="<%=rsFinancialEligibility.Fields.Item("fltGrantAmt_Tech").Value%>" onChange="ChangeAmount();" size="9" tabindex="18" onKeypress="AllowNumericOnly();">
			<span style="font-size: 7pt">(Estimated Amount Used For Technology)</span>
		</td>
	</tr>
	<tr>
		<td nowrap></td>
		<td nowrap>
			$<input type="text" name="AmountUsedForServices" value="<%=rsFinancialEligibility.Fields.Item("fltGrantAmt_Srv").Value%>" onChange="ChangeAmount();" size="9" tabindex="19" onKeypress="AllowNumericOnly();">
			<span style="font-size: 7pt">(Estimated Amount Used For Services)</span>
		</td>
	</tr>
	<tr>
		<td nowrap>Grant Amount Left:</td>
		<td nowrap><input type="text" name="GrantAmountLeft" value="$0.00" size="12" tabindex="20" onKeypress="AllowNumericOnly();"></td>
	</tr>
	<tr>
		<td nowrap valign="top">Comments:</td>
		<td nowrap valign="top"><textarea name="Comments" tabindex="21" accesskey="L" cols="60" rows="4"><%=rsFinancialEligibility.Fields.Item("chvNote").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="22" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="23" onClick="history.go(-1);" class="btnstyle"></td>
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