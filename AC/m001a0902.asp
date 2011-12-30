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
	var temp = String(Request.Form("CC")).split(":");
	var CC = new Array(10);
	for (var i = 0; i < 10; i++) CC[i] = 0;
	for (var i = 0; i < temp.length; i++) CC[i] = temp[i];
	if (temp == "") CC[0] = 0;
	var Is_Recipient_Client = 1;
	var temp2 = String(Request.Form("Recipient")).split(":");
	Is_Recipient_Client = ((temp2[0]=="Client")?1:0);
	var DocumentName = String(Request.Form("DocumentName")).replace(/'/g, "''");	

	CC[9] = ((Request.Form("CCClient")=="on")?"99999":"0");
	
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
	switch(String(Request.Form("Template"))) {
		//CSG Accept
		case "861":
			var DocumentConditionOther = String(Request.Form("DocumentConditionOther")).replace(/'/g, "''");	
			var Conditions = String(Request.Form("Conditions")).replace(/'/g, "''");	
			var Donation = String(Request.Form("Donation")).replace(/'/g, "''");	
			var ConfigurationRequested = String(Request.Form("ConfigurationRequested")).replace(/'/g, "''");	
			var LoanReturn = String(Request.Form("LoanReturn")).replace(/'/g, "''");	
			var OtherCSGMIRIssue = String(Request.Form("OtherCSGMIRIssue")).replace(/'/g, "''");				
			var TrainingRequested = ((Request.Form("TrainingRequested")=="on")?"1":"0");
			rsTemplate.Source = "{call dbo.cp_insert_crspltr_csg_accept("+Request.QueryString("intBuyout_Req_id")+","+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("DocumentConditionOne")+","+Request.Form("DocumentConditionTwo")+","+Request.Form("DocumentConditionThree")+","+Request.Form("DocumentConditionFour")+",'"+DocumentConditionOther+"','"+Conditions+"','"+Donation+"','"+ConfigurationRequested+"','"+LoanReturn+"',"+Request.Form("ShippingOrigin")+","+TrainingRequested+",0)}";
		break;
	}
	rsTemplate.CursorType = 0;
	rsTemplate.CursorLocation = 2;
	rsTemplate.LockType = 3;
	rsTemplate.Open();	
	Response.Redirect("../BO/m010FS01.asp?intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));
}

var rsContact = Server.CreateObject("ADODB.Recordset");
rsContact.ActiveConnection = MM_cnnASP02_STRING;
rsContact.Source = "{call dbo.cp_ClnCtact("+ Request.QueryString("intAdult_id") + ")}";
rsContact.CursorType = 0;
rsContact.CursorLocation = 2;
rsContact.LockType = 3;
rsContact.Open();

var rsClient = Server.CreateObject("ADODB.Recordset");
rsClient.ActiveConnection = MM_cnnASP02_STRING;
rsClient.Source = "{call dbo.cp_Idv_Adult_Client(" + Request.QueryString("intAdult_id") + ")}";
rsClient.CursorType = 0;
rsClient.CursorLocation = 2;
rsClient.LockType = 3;
rsClient.Open();

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

var rsDocumentConditions = Server.CreateObject("ADODB.Recordset");
rsDocumentConditions.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentConditions.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,14,'',2,'Q',0)}";
rsDocumentConditions.CursorType = 0;
rsDocumentConditions.CursorLocation = 2;
rsDocumentConditions.LockType = 3;
rsDocumentConditions.Open();
%>									
									
<html>
<head>
	<title>New Correspondence</title>
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
	var DocumentArray = new Array(<%=count%>);
<% 
var i = 0;
while (!rsTemplate.EOF) {
	if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m010") {	
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
				
	function Save(){
		if (!CheckDate(document.frm0902.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0902.DateGenerated.focus();
			return ;
		}
		if (Trim(document.frm0902.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0902.DocumentName.focus();
			return ;
		}
			
		document.frm0902.CC.value = "";
		for (var i = 0; i < document.frm0902.CCList.options.length; i++) {
			if (document.frm0902.CCList.options[i].selected) {
				document.frm0902.CC.value = document.frm0902.CC.value + ":" + document.frm0902.CCList.options[i].value;
			}
		}
		if (document.frm0902.CC.value.length > 0) document.frm0902.CC.value = document.frm0902.CC.value.substring(1, document.frm0902.CC.value.length);

		var temp = document.frm0902.action;
				
		if (document.frm0902.MailMethod.value=="0") {				
			if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();			
			document.frm0902.action = "../TPL/"+DocumentArray[document.frm0902.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
		} else {
			document.frm0902.action = "../TPL/E-"+DocumentArray[document.frm0902.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
		}
							
		document.frm0902.target = "_blank";
		document.frm0902.submit();

		document.frm0902.action = temp;
		document.frm0902.target = "_self";
		document.frm0902.submit();				
	}

	function GenerateEnvelope(){
		//Print recipient
		var temp = document.frm0901.Recipient.value.split(":");
		document.frm0901.action = "../TPL/PrintEnvelope.asp?RecipientType=" + temp[0] + "&To=" + temp[1];		
		document.frm0901.target = "_blank";
		document.frm0901.submit();

		document.frm0901.CC.value = "";
		for (var i = 0; i < document.frm0901.CCList.options.length; i++) {
			if (document.frm0901.CCList.options[i].selected) document.frm0901.CC.value = document.frm0901.CC.value + ":" + document.frm0901.CCList.options[i].value;
		}
		
		if (document.frm0901.CC.value.length > 0) document.frm0901.CC.value = document.frm0901.CC.value.substring(1, document.frm0901.CC.value.length);
		
		//Print CCs
		for (var i = 0; i < document.frm0901.CCList.options.length; i++) {
			if (document.frm0901.CCList.options[i].selected) {
				document.frm0901.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0901.CCList.options[i].value;
				document.frm0901.target = "_blank";
				document.frm0901.submit();
			}
		}
		
		//if CC Client
		if (document.frm0901.CCClient.checked) {
			document.frm0901.action = "../TPL/PrintEnvelope.asp?RecipientType=Client&To=<%=Request.QueryString("intAdult_id")%>";
			document.frm0901.target = "_blank";
			document.frm0901.submit();
		}		
	}

	function ChangeType(){
		if (document.frm0902.Type.value == "0") {
			window.location.href = "m001a0904.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
		}
	}	
	</script>
</head>
<body onLoad="document.frm0902.Subject.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0902">
<h5>New Correspondence</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><select name="Subject" tabindex="1" accesskey="F">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsClient.Fields.Item("intAdult_Id").Value == "Request.QueryString(\"intAdult_id\")")?"SELECTED":"")%> ><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		rsClient.MoveFirst();
		%>
		</select></td>
    </tr>
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="2" onChange="ChangeType();">
			<option value="4" SELECTED>Form Letter
			<option value="0">Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="3">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="Client:<%=(rsClient.Fields.Item("intAdult_Id").Value)%>"><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		%>		
		<% 
		while (!rsContact.EOF) {
		%>
			<option value="Contact:<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
		<%
			rsContact.MoveNext();
		}
		rsContact.Requery();
		%>		
		</select></td>
	</tr>
	<tr>
		<td valign="top">CC:</td>
		<td valign="top"><select name="CCList" multiple size="5" tabindex="4">
		<% 
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>"><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
		<%
			rsContact.MoveNext();
		}
		%>		
		</select></td>		
	</tr>
	<tr>
		<td></td>	
		<td><input type="checkbox" name="CCClient" tabindex="5" class="chkstyle">CC the client</td>
	</tr>	
    <tr> 
		<td nowrap>Template:</td>
		<td nowrap><select name="Template" tabindex="6">
	<% 
	while (!rsTemplate.EOF) {
		if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m010") {
	%>
			<option value="<%=(rsTemplate.Fields.Item("insTemplate_id").Value)%>"><%=(rsTemplate.Fields.Item("chvTemplate_Name").Value)%></option>
	<%
		}
		rsTemplate.MoveNext();
	}
	%>
        </select></td>
	</tr>
    <tr> 
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" maxlength="50" size="30" tabindex="7"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="9" accesskey="L">
			<option value="0">Canada Post
			<option value="1">E-Mail
		</select></td>
	</tr>	
</table>
<hr>
<h5>CSG Accept</h5>
<b>Technology Plan</b>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Equipment Conditions:</td>
		<td><input type="text" name="Conditions" maxlength="80" size="65" tabindex="10">&nbsp;(1)</td>
	</tr>
	<tr>
		<td>Donation:</td>
		<td><input type="text" name="Donation" maxlength="80" size="65" tabindex="11">&nbsp;(2)</td>
	</tr>
	<tr>
		<td>Configuration Requested:</td>
		<td><input type="text" name="ConfigurationRequested" maxlength="80" size="65" tabindex="12">&nbsp;(3)</td>
	</tr>
	<tr>
		<td>Loan Return:</td>
		<td><input type="text" name="LoanReturn" maxlength="80"size="65" tabindex="13">&nbsp;(4)</td>
	</tr>
	<tr>
		<td>Shipping Origin:</td>
		<td><select name="ShippingOrigin" tabindex="14">
				<option value="1">ASP
				<option value="2">Vendors
				<option value="3">Both
		</select>&nbsp;(5)</td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="TrainingRequested" class="chkstyle" tabindex="15">Training Requested&nbsp;(6)</td>
	</tr>
</table>
<br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap><b>Documentation Conditions:</b></td>
		<td><select name="DocumentConditionOne" tabindex="16">
			<option value="0">(none)
		<%
		rsDocumentConditions.MoveFirst();
		while (!rsDocumentConditions.EOF) {
		%>
			<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>"><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
		<%
			rsDocumentConditions.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionTwo" tabindex="17">
			<option value="0">(none)
		<%
		rsDocumentConditions.MoveFirst();
		while (!rsDocumentConditions.EOF) {
		%>
			<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>"><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
		<%
			rsDocumentConditions.MoveNext();
		}
		%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionThree" tabindex="18">
				<option value="0">(none)
			<%
			rsDocumentConditions.MoveFirst();
			while (!rsDocumentConditions.EOF) {
			%>
				<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>"><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
			<%
				rsDocumentConditions.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionFour" tabindex="19">
				<option value="0">(none)
			<%
			rsDocumentConditions.MoveFirst();
			while (!rsDocumentConditions.EOF) {
			%>
				<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>"><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
			<%
				rsDocumentConditions.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="DocumentConditionOther" size="65" maxlength="80" tabindex="20">&nbsp;(11)</td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Generate Letter" onClick="Save();" class="btnstyle" tabindex="21"></td>
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle" tabindex="22"></td>
    </tr>
</table>
<input type="hidden" name="CC" value="">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
%>