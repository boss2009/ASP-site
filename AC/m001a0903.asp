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

	CC[9] = ((Request.Form("CCClient")=="on")?"99999":"0");
	
	var DocumentName = String(Request.Form("DocumentName")).replace(/'/g, "''");	
	var LoanConditionOther = String(Request.Form("LoanConditionOther")).replace(/'/g, "''");	
	var DocumentConditionOther = String(Request.Form("DocumentConditionOther")).replace(/'/g, "''");
	var TrainingRequested = ((Request.Form("TrainingRequested")=="on")?"1":"0");
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
	rsTemplate.Source = "{call dbo.cp_insert_crspltr_loan_accept("+Request.QueryString("intLoan_req_id")+","+Request.QueryString("intAdult_id")+","+Session("insStaff_id")+","+Request.Form("Subject")+",0,"+Request.Form("Recipient")+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+","+Request.Form("Template")+",'"+DocumentName+"',1,'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+","+Request.Form("LoanConditionOne")+","+Request.Form("LoanConditionTwo")+","+Request.Form("LoanConditionThree")+","+Request.Form("LoanConditionFour")+",'"+LoanConditionOther+"',"+Request.Form("DocumentConditionOne")+","+Request.Form("DocumentConditionTwo")+","+Request.Form("DocumentConditionThree")+","+Request.Form("DocumentConditionFour")+",'"+DocumentConditionOther+"',"+TrainingRequested+",0)}";
	rsTemplate.CursorType = 0;
	rsTemplate.CursorLocation = 2;
	rsTemplate.LockType = 3;
	rsTemplate.Open();
	Response.Redirect("../LN/m008FS01.asp?intLoan_req_id="+Request.QueryString("intLoan_req_id"));
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

var rsLoanConditions = Server.CreateObject("ADODB.Recordset");
rsLoanConditions.ActiveConnection = MM_cnnASP02_STRING;
rsLoanConditions.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,15,'',2,'Q',0)}";
rsLoanConditions.CursorType = 0;
rsLoanConditions.CursorLocation = 2;
rsLoanConditions.LockType = 3;
rsLoanConditions.Open();

var rsDocumentConditions = Server.CreateObject("ADODB.Recordset");
rsDocumentConditions.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentConditions.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,22,'',2,'Q',0)}";
rsDocumentConditions.CursorType = 0;
rsDocumentConditions.CursorLocation = 2;
rsDocumentConditions.LockType = 3;
rsDocumentConditions.Open();
%>																																
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
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
	if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m008") {			
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
		document.frm0903.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0903.Recipient.value;		
		document.frm0903.target = "_blank";
		document.frm0903.submit();

		document.frm0903.CC.value = "";
		
		for (var i = 0; i < document.frm0903.CCList.options.length; i++) {
			if (document.frm0903.CCList.options[i].selected) {
				document.frm0903.CC.value = document.frm0903.CC.value + ":" + document.frm0903.CCList.options[i].value;
			}
		}
		
		if (document.frm0903.CC.value.length > 0) {
			document.frm0903.CC.value = document.frm0903.CC.value.substring(1, document.frm0903.CC.value.length);
		}
		
		//Print CCs
		for (var i = 0; i < document.frm0903.CCList.options.length; i++) {
			if (document.frm0903.CCList.options[i].selected) {
				document.frm0903.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0903.CCList.options[i].value;
				document.frm0903.target = "_blank";
				document.frm0903.submit();
			}
		}
		
		//if CC client	
		if (document.frm0903.CCClient.checked) {
			document.frm0903.action = "../TPL/PrintEnvelope.asp?RecipientType=Client&To=<%=Request.QueryString("intAdult_id")%>";
			document.frm0903.target = "_blank";
			document.frm0903.submit();
		}	
	}
			
	function Save(){
		if (!CheckDate(document.frm0903.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0903.DateGenerated.focus();
			return ;
		}
		
		if (Trim(document.frm0903.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0903.DocumentName.focus();
			return ;
		}
				
		document.frm0903.CC.value = "";
		for (var i = 0; i < document.frm0903.CCList.options.length; i++) {
			if (document.frm0903.CCList.options[i].selected) {
				document.frm0903.CC.value = document.frm0903.CC.value + ":" + document.frm0903.CCList.options[i].value;
			}
		}
		
		if (document.frm0903.CC.value.length > 0) {
			document.frm0903.CC.value = document.frm0903.CC.value.substring(1, document.frm0903.CC.value.length);
		}

		var temp = document.frm0903.action;
		
		if (document.frm0903.MailMethod.value=="0") {
			if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();					
			document.frm0903.action = "../TPL/"+DocumentArray[document.frm0903.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";
		} else {
			document.frm0903.action = "../TPL/E-"+DocumentArray[document.frm0903.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";
		}			
		document.frm0903.target = "_blank";
		document.frm0903.submit();
		
		document.frm0903.action = temp;
		document.frm0903.target = "_self";
		document.frm0903.submit();
	}
	
	function ChangeType(){
		if (document.frm0903.Type.value == "0") {
			window.location.href = "m001a0904.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";
		}
	}	
	</script>
</head>
<body onLoad="document.frm0903.Subject.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0903">
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
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=((Trim(rsContact.Fields.Item("chvRelationship").Value)=="Referring Agent")?"SELECTED":"")%>><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
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
    <tr> 
		<td nowrap>Template:</td>
		<td nowrap><select name="Template" tabindex="6">
	<% 
	while (!rsTemplate.EOF) {
		if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m008") {		
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
<h5>Loan Accept</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Loan Conditions:</b></td>
		<td><select name="LoanConditionOne" tabindex="10">
				<option value="0">(none)		
			<%
			rsLoanConditions.MoveFirst();
			while (!rsLoanConditions.EOF) {
			%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>"><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanConditions.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanConditionTwo" tabindex="11">
				<option value="0">(none)
			<%
			rsLoanConditions.MoveFirst();
			while (!rsLoanConditions.EOF) {
			%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>"><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanConditions.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanConditionThree" tabindex="12">
				<option value="0">(none)
			<%
			rsLoanConditions.MoveFirst();
			while (!rsLoanConditions.EOF) {
			%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>"><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanConditions.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanConditionFour" tabindex="13">
				<option value="0">(none)
			<%
			rsLoanConditions.MoveFirst();
			while (!rsLoanConditions.EOF) {
			%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>"><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
			<%
				rsLoanConditions.MoveNext();
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="LoanConditionOther" maxlength="80" size="65" tabindex="14">&nbsp;(4)</td>
	</tr>
</table>
<br><br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Documentation Conditions:</b></td>
		<td><select name="DocumentConditionOne" tabindex="15">
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
		<td><select name="DocumentConditionTwo" tabindex="16">
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
		<td><select name="DocumentConditionThree" tabindex="17">
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
		<td><select name="DocumentConditionFour" tabindex="18">
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
		<td><input type="text" name="DocumentConditionOther" maxlength="80" size="65" tabindex="19"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="TrainingRequested" class="chkstyle" tabindex="20"><b>Training Requested</b></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Generate Letter" onClick="Save();" tabindex="21" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="22" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="CC">
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
%>