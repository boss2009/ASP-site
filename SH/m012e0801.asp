<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

var rsLetter = Server.CreateObject("ADODB.Recordset");
rsLetter.ActiveConnection = MM_cnnASP02_STRING;
rsLetter.Source = "{call dbo.cp_get_ac_crsp_hstry3("+ Request.QueryString("intLetter_id") + ",0)}";
rsLetter.CursorType = 0;
rsLetter.CursorLocation = 2;
rsLetter.LockType = 3;
rsLetter.Open();

var rsCC = Server.CreateObject("ADODB.Recordset");
rsCC.ActiveConnection = MM_cnnASP02_STRING;
rsCC.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,1,'Q',0)}";
rsCC.CursorType = 0;
rsCC.CursorLocation = 2;
rsCC.LockType = 3;
rsCC.Open();

var rsCondition = Server.CreateObject("ADODB.Recordset");
rsCondition.ActiveConnection = MM_cnnASP02_STRING;
rsCondition.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,2,'Q',0)}";
rsCondition.CursorType = 0;
rsCondition.CursorLocation = 2;
rsCondition.LockType = 3;
rsCondition.Open();

var rsReason = Server.CreateObject("ADODB.Recordset");
rsReason.ActiveConnection = MM_cnnASP02_STRING;
rsReason.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,4,'Q',0)}";
rsReason.CursorType = 0;
rsReason.CursorLocation = 2;
rsReason.LockType = 3;
rsReason.Open();

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_school_contacts("+ Request.QueryString("insSchool_id") + ",0,0,0,'Q',0)}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsInstitution = Server.CreateObject("ADODB.Recordset");
rsInstitution.ActiveConnection = MM_cnnASP02_STRING;
rsInstitution.Source = "{call dbo.cp_school2("+Request.QueryString("insSchool_id")+",'',0,0,0,0,0,0,0,'',1,'Q',0)}";
rsInstitution.CursorType = 0;
rsInstitution.CursorLocation = 2;
rsInstitution.LockType = 3;
rsInstitution.Open();

var rsTemplate = Server.CreateObject("ADODB.Recordset");
rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
rsTemplate.Source = "{call dbo.cp_Letter_template(0,1,'',0,'',0,0,0,0,0,0,2,'Q',0)}";
rsTemplate.CursorType = 0;
rsTemplate.CursorLocation = 2;
rsTemplate.LockType = 3;
rsTemplate.Open();
var count = 0;
while (!rsTemplate.EOF) {
	count++;
	rsTemplate.MoveNext();
}
rsTemplate.MoveFirst();

var rsDeclineReason = Server.CreateObject("ADODB.Recordset");
rsDeclineReason.ActiveConnection = MM_cnnASP02_STRING;
rsDeclineReason.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,7,'',2,'Q',0)}";
rsDeclineReason.CursorType = 0;
rsDeclineReason.CursorLocation = 2;
rsDeclineReason.LockType = 3;
rsDeclineReason.Open();
%>									
									
<html>
<head>
	<title>View Correspondence</title>
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
				window.location.href="m012q0801.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
			break;
		}
	}
	</script>	
	<script language="Javascript">
	var DocumentArray = new Array(<%=count%>);
<% 
var i = 0;
while (!rsTemplate.EOF) {
	if ((rsTemplate.Fields.Item("chvFileName").Value=="m012tpl001.asp") || (rsTemplate.Fields.Item("chvFileName").Value=="m012tpl003.asp")) {	
%>
		DocumentArray[<%=i%>] = new Array(3);
		DocumentArray[<%=i%>][0] = <%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>;
		DocumentArray[<%=i%>][1] = "<%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%>";
		DocumentArray[<%=i%>][2] = "<%=(rsTemplate.Fields.Item("chvFileName").Value)%>";
<%
		i++;
	}
	rsTemplate.MoveNext();
}
rsTemplate.MoveFirst();
%>

	function GenerateEnvelope(){
		//Print recipient
		document.frm0801.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0801.Recipient.value;
		document.frm0801.target = "_blank";
		document.frm0801.submit();

		document.frm0801.CC.value = "";
		for (var i = 0; i < document.frm0801.CCList.options.length; i++) {
			if (document.frm0801.CCList.options[i].selected) {
				document.frm0801.CC.value = document.frm0801.CC.value + ":" + document.frm0801.CCList.options[i].value;
			}
		}
		
		if (document.frm0801.CC.value.length > 0) {
			document.frm0801.CC.value = document.frm0801.CC.value.substring(1, document.frm0801.CC.value.length);
		}
		
		//Print CCs
		for (var i = 0; i < document.frm0801.CCList.options.length; i++) {
			if (document.frm0801.CCList.options[i].selected) {
				document.frm0801.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0801.CCList.options[i].value;
				document.frm0801.target = "_blank";
				document.frm0801.submit();
			}
		}		
	}
		
	function Save(){
		if (!CheckDate(document.frm0801.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0801.DateGenerated.focus();
			return ;
		}
		
		if (Trim(document.frm0801.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0801.DocumentName.focus();
			return ;
		}

		if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();
		
		document.frm0801.CC.value = "";
		for (var i = 0; i < document.frm0801.CCList.options.length; i++) {
			if (document.frm0801.CCList.options[i].selected) {
				document.frm0801.CC.value = document.frm0801.CC.value + ":" + document.frm0801.CCList.options[i].value;
			}
		}
		if (document.frm0801.CC.value.length > 0) {
			document.frm0801.CC.value = document.frm0801.CC.value.substring(1, document.frm0801.CC.value.length);
		}
		
		if (document.frm0801.MailMethod.value=="0") {				
			document.frm0801.action = "../TPL/"+DocumentArray[document.frm0801.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
		} else {
			document.frm0801.action = "../TPL/E-"+DocumentArray[document.frm0801.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
		}
		
		document.frm0801.target = "_blank";
		document.frm0801.submit();
	}
	
	function ChangeType(){
		if (document.frm0801.Type.value == "0") {
			window.location.href = "m012a0803.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>";
		}
	}

	function ChangeTemplate(){
		PILATAccept.style.visibility = "hidden";
		PILATDecline.style.visibility = "hidden";
		switch (String(document.frm0801.Template.value)) {
			case "870":
				PILATAccept.style.visibility = "visible";			
				buttons.style.top = "410px";
			break;
			case "867":
				PILATDecline.style.visibility = "visible";
				buttons.style.top = "470px";
			break;
		}
	}
	
	function Init(){
		document.frm0801.Type.focus();
		ChangeTemplate();
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0801">
<h5>View Correspondence</h5>
<i>This page is readonly.</i>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="1" accesskey="F">
			<option value="4" <%=((rsLetter.Fields.Item("chvRx_Type").Value=="Form Letter")?"SELECTED":"")%>>Form Letter
			<option value="0" <%=((rsLetter.Fields.Item("chvRx_Type").Value=="Custom Letter")?"SELECTED":"")%>>Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="2">
		<%
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=(((rsLetter.Fields.Item("chvRx_Class").Value=="Contact")&&(rsContact.Fields.Item("intContact_id").Value==rsLetter.Fields.Item("intRecipient_id").Value))?"SELECTED":"")%>><%=rsContact.Fields.Item("chvFst_Name").Value%>&nbsp;<%=rsContact.Fields.Item("chvLst_Name").Value%> - <%=(rsContact.Fields.Item("chvRelationship").Value)%>
		<%
			rsContact.MoveNext();
		}
		rsContact.Requery();		
		%>		
		</select></td>
	</tr>
	<tr>
		<td valign="top">CC:</td>
		<td valign="top"><select name="CCList" multiple size="5" tabindex="3">
		<% 
		var selected;
		while (!rsContact.EOF) {
			rsCC.Requery();
			selected = "";
			while (!rsCC.EOF) {
				if (rsContact.Fields.Item("intContact_id").Value==rsCC.Fields.Item("intContact_Id").Value) selected = "SELECTED";
				rsCC.MoveNext();
			}		
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=selected%>><%=rsContact.Fields.Item("chvFst_Name").Value%>&nbsp;<%=rsContact.Fields.Item("chvLst_Name").Value%> - <%=(rsContact.Fields.Item("chvRelationship").Value)%>
		<%			
			rsContact.MoveNext();
		}
		%>		
		</select></td>
	</tr>
    <tr> 
		<td>Template:</td>
		<td><select name="Template" tabindex="4">
	<% 
	while (!rsTemplate.EOF) {
		if ((rsTemplate.Fields.Item("chvFileName").Value=="m012tpl001.asp") || (rsTemplate.Fields.Item("chvFileName").Value=="m012tpl003.asp")) {
	%>
			<option value="<%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>" <%=((rsTemplate.Fields.Item("insTemplate_id").Value==Request.QueryString("insTemplate_id"))?"SELECTED":"")%>><%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%></option>
	<%
		}
		rsTemplate.MoveNext();
	}
	%>
        </select></td>
	</tr>
    <tr> 
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" value="<%=rsLetter.Fields.Item("chvLetter_Name").Value%>" maxlength="50" size="30" tabindex="5"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=FilterDate(rsLetter.Fields.Item("dtsSend_Date").Value)%>" size="11" maxlength="10" tabindex="6" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="7" accesskey="L">
			<option value="0" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="Canada_Post")?"SELECTED":"")%>>Canada Post
			<option value="1" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="e-Mail")?"SELECTED":"")%>>E-Mail
		</select></td>
	</tr>	
</table>
<hr>
<div id="PILATAccept" style="position: absolute; top: 305px">
<h5>PILAT Accept Consult/Training</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Type of Referral:</b></td>
		<td><select name="PILATAcceptReferralType">
				<option value="5" <%=((rsLetter.Fields.Item("insINT_01").Value==5)?"SELECTED":"")%>>Training Only
				<option value="6" <%=((rsLetter.Fields.Item("insINT_01").Value==6)?"SELECTED":"")%>>Consultation Only
		</select></td>
	</tr>
	<tr>
		<td><b>Conditions:</b></td>
		<td><select name="Conditions">
				<option value="1" <%if (!rsCondition.EOF) Response.Write((rsCondition.Fields.Item("intDoc_Id").Value==1)?"SELECTED":"")%>>Training Only Conditions
				<option value="2" <%if (!rsCondition.EOF) Response.Write((rsCondition.Fields.Item("intDoc_Id").Value==2)?"SELECTED":"")%>>Consultation Only Conditions
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherConditions" value="<%=rsLetter.Fields.Item("chvText_01").Value%>" size="60" maxlength="80"></td>
	</tr>
</table>
</div>
<div id="PILATDecline" style="position: absolute; top: 305px">
<h5>PILAT Decline</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Type of Referral:</b></td>
		<td><select name="PilatDeclineReferralType">
				<option value="1" <%=((rsLetter.Fields.Item("insINT_01").Value==1)?"SELECTED":"")%>>Low Utilization
				<option value="2" <%=((rsLetter.Fields.Item("insINT_01").Value==2)?"SELECTED":"")%>>Interim
				<option value="3" <%=((rsLetter.Fields.Item("insINT_01").Value==3)?"SELECTED":"")%>>Donation
				<option value="4" <%=((rsLetter.Fields.Item("insINT_01").Value==4)?"SELECTED":"")%>>Purchase
		</select></td>
	</tr>
	<tr>
		<td><b>Reasons:</b></td>
		<td><select name="DeclineReasonOne">
				<option value="0">(none)
	<%			
	rsReason.Requery();	
	if (!rsReason.EOF) {
		while (!rsDeclineReason.EOF) {
			if (rsDeclineReason.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {
	%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDeclineReason.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DeclineReasonTwo">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsDeclineReason.Requery();	
		while (!rsDeclineReason.EOF) {
			if (rsDeclineReason.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDeclineReason.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DeclineReasonThree">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsDeclineReason.Requery();	
		while (!rsDeclineReason.EOF) {
			if (rsDeclineReason.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDeclineReason.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DeclineReasonFour">
				<option value="0">(none)
	<%
	if (!rsReason.EOF) {
		rsDeclineReason.Requery();	
		while (!rsDeclineReason.EOF) {
			if (rsDeclineReason.Fields.Item("intDoc_id").Value==rsReason.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDeclineReason.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDeclineReason.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDeclineReason.MoveNext();
		}
		rsReason.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherDeclineReason" value="<%=rsLetter.Fields.Item("chvText_01").Value%>" maxlength="80" size="60"></td>		
	</tr>
</table>
</div>
<div id="buttons" style="position: absolute; top: 450px">
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="View Letter" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m012q0801.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';" class="btnstyle"></td>
    </tr>
</table>
</div>
<input type="hidden" name="CC" value="">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsInstitution.Close();
rsContact.Close();
rsTemplate.Close();
%>