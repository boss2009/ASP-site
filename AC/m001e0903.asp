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

var CCClient = false;
while ((!rsCC.EOF) && (CCClient==false)) {
	if (rsCC.Fields.Item("intContact_Id").Value==99999) CCClient = true;
	rsCC.MoveNext();
}
rsCC.Requery();

var rsCondition = Server.CreateObject("ADODB.Recordset");
rsCondition.ActiveConnection = MM_cnnASP02_STRING;
rsCondition.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,2,'Q',0)}";
rsCondition.CursorType = 0;
rsCondition.CursorLocation = 2;
rsCondition.LockType = 3;
rsCondition.Open();

var rsPurpose = Server.CreateObject("ADODB.Recordset");
rsPurpose.ActiveConnection = MM_cnnASP02_STRING;
rsPurpose.Source = "{call dbo.cp_crsp_ltr_assc("+Request.QueryString("intLetter_id")+",0,3,'Q',0)}";
rsPurpose.CursorType = 0;
rsPurpose.CursorLocation = 2;
rsPurpose.LockType = 3;
rsPurpose.Open();

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
				window.location.href="m001q0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
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
		
		//if CC Client
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
		
		if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();	
				
		document.frm0903.CC.value = "";
		for (var i = 0; i < document.frm0903.CCList.options.length; i++) {
			if (document.frm0903.CCList.options[i].selected) {
				document.frm0903.CC.value = document.frm0903.CC.value + ":" + document.frm0903.CCList.options[i].value;
			}
		}
		
		if (document.frm0903.CC.value.length > 0) {
			document.frm0903.CC.value = document.frm0903.CC.value.substring(1, document.frm0903.CC.value.length);
		}
	
		if (document.frm0903.MailMethod.value=="0") {		
			document.frm0903.action = "../TPL/"+DocumentArray[document.frm0903.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";
		} else {
			document.frm0903.action = "../TPL/E-"+DocumentArray[document.frm0903.Template.selectedIndex][2]+"?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";
		}			
		document.frm0903.target = "_blank";
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
<h5>View Correspondence</h5>
<i>This page is readonly.</i>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><select name="Subject" tabindex="1" accesskey="F">
		<% 
		while (!rsClient.EOF) {
		%>
			<option value="<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsClient.Fields.Item("intAdult_Id").Value == rsLetter.Fields.Item("intSubject_id").Value)?"SELECTED":"")%>><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		%>  
		</select></td>
    </tr>
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="2">
			<option value="4" <%=((rsLetter.Fields.Item("chvRx_Type").Value=="Form Letter")?"SELECTED":"")%>>Form Letter
			<option value="0" <%=((rsLetter.Fields.Item("chvRx_Type").Value=="Custom Letter")?"SELECTED":"")%>>Custom Letter
		</select></td> 
	</tr>
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="3">
		<% 
		while (!rsContact.EOF) {
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=(((rsLetter.Fields.Item("chvRx_Class").Value=="Contact")&&(rsContact.Fields.Item("intContact_id").Value==rsLetter.Fields.Item("intRecipient_id").Value))?"SELECTED":"")%>><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
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
		var selected;
		while (!rsContact.EOF) {
			rsCC.Requery();
			selected = "";
			while (!rsCC.EOF) {
				if (rsContact.Fields.Item("intContact_id").Value==rsCC.Fields.Item("intContact_Id").Value) selected = "SELECTED";
				rsCC.MoveNext();
			}		
		%>
			<option value="<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=selected%>><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)		
		<%			
			rsContact.MoveNext();
		}
		%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><input type="checkbox" name="CCClient" tabindex="5" <%=((CCClient)?"CHECKED":"")%> class="chkstyle">CC the client</td>
	</tr>		
    <tr> 
		<td nowrap>Template:</td>
		<td nowrap><select name="Template" tabindex="6">
	<% 
	while (!rsTemplate.EOF) {
		if (rsTemplate.Fields.Item("chvFileName").Value.substring(0,4)=="m008") {		
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
		<td nowrap><input type="text" name="DocumentName" value="<%=rsLetter.Fields.Item("chvLetter_Name").Value%>" maxlength="50" size="30" tabindex="7"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=FilterDate(rsLetter.Fields.Item("dtsSend_Date").Value)%>" size="11" maxlength="10" tabindex="8" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="9" accesskey="L">
			<option value="0" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="Canada_Post")?"SELECTED":"")%>>Canada Post
			<option value="1" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="e-Mail")?"SELECTED":"")%>>E-Mail
		</select></td>
	</tr>		
</table>
<hr>
<h5>Loan Accept</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Loan Conditions:</b></td>
		<td><select name="LoanConditionOne">
				<option value="0">(none)		
	<%	
	rsCondition.Requery();	
	if (!rsCondition.EOF) {
		while (!rsLoanConditions.EOF) {
			if (rsLoanConditions.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {
	%>	
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsLoanConditions.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanConditionTwo">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsLoanConditions.Requery();	
		while (!rsLoanConditions.EOF) {
			if (rsLoanConditions.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsLoanConditions.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanConditionThree">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsLoanConditions.Requery();	
		while (!rsLoanConditions.EOF) {
			if (rsLoanConditions.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsLoanConditions.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="LoanConditionFour">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsLoanConditions.Requery();	
		while (!rsLoanConditions.EOF) {
			if (rsLoanConditions.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsLoanConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsLoanConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsLoanConditions.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>		
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="LoanConditionOther" value="<%=Trim(rsLetter.Fields.Item("chvText_01").Value)%>" maxlength="80" size="65">&nbsp;(4)</td>
	</tr>
</table>
<br><br>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Documentation Conditions:</b></td>
		<td><select name="DocumentConditionOne">
				<option value="0">(none)
	<%	
	rsPurpose.Requery();	
	if (!rsPurpose.EOF) {
		while (!rsDocumentConditions.EOF) {
			if (rsDocumentConditions.Fields.Item("intDoc_id").Value==rsPurpose.Fields.Item("intDoc_Id").Value) {
	%>				
				<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDocumentConditions.MoveNext();
		}
		rsPurpose.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionTwo">
				<option value="0">(none)
	<%
	if (!rsPurpose.EOF) {
		rsDocumentConditions.Requery();	
		while (!rsDocumentConditions.EOF) {
			if (rsDocumentConditions.Fields.Item("intDoc_id").Value==rsPurpose.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDocumentConditions.MoveNext();
		}
		rsPurpose.MoveNext();
	}
	%>		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionThree">
				<option value="0">(none)
	<%
	if (!rsPurpose.EOF) {
		rsDocumentConditions.Requery();	
		while (!rsDocumentConditions.EOF) {
			if (rsDocumentConditions.Fields.Item("intDoc_id").Value==rsPurpose.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDocumentConditions.MoveNext();
		}
		rsPurpose.MoveNext();
	}
	%>
			</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionFour">
				<option value="0">(none)
	<%
	if (!rsPurpose.EOF) {
		rsDocumentConditions.Requery();	
		while (!rsDocumentConditions.EOF) {
			if (rsDocumentConditions.Fields.Item("intDoc_id").Value==rsPurpose.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDocumentConditions.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentConditions.Fields.Item("chvDocDesc").Value%>
	<%
			}
			rsDocumentConditions.MoveNext();
		}
		rsPurpose.MoveNext();
	}
	%>		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="DocumentConditionOther" value="<%=Trim(rsLetter.Fields.Item("chvText_02").Value)%>" maxlength="80" size="65"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="TrainingRequested" <%=((rsLetter.Fields.Item("bitCkBx_01").Value=="1")?"CHECKED":"")%> class="chkstyle"><b>Training Requested</b></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="View Letter" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m001q0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>';" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="CC">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
rsTemplate.Close();
rsLetter.Close();
rsCC.Close();
rsPurpose.Close();
rsCondition.Close();
%>