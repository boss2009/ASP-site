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
				window.location.href="m001q0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
			break;			
		}
	}
	</script>	
	<script language="Javascript">	
	function Save(){
		if (!CheckDate(document.frm0904.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0904.DateGenerated.focus();
			return ;
		}
		
		if (Trim(document.frm0904.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0904.DocumentName.focus();
			return ;
		}
		
		if (document.frm0904.CustomLetterContent.value.length > 4000) {
			alert("Custom letter content cannot exceed 4000 characters.");
			document.frm0904.CustomLetterContent.focus();
			return ;
		}

		if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();	
		
		document.frm0904.CC.value = "";
		
		for (var i = 0; i < document.frm0904.CCList.options.length; i++) {
			if (document.frm0904.CCList.options[i].selected) document.frm0904.CC.value = document.frm0904.CC.value + ":" + document.frm0904.CCList.options[i].value;
		}
		if (document.frm0904.CC.value.length > 0) document.frm0904.CC.value = document.frm0904.CC.value.substring(1, document.frm0904.CC.value.length);
						
		if (document.frm0904.MailMethod.value=="0") {		
			document.frm0904.action = "../TPL/CustomLetterTemplate.asp";
		} else {
			document.frm0904.action = "../TPL/E-CustomLetterTemplate.asp";
		}		

		document.frm0904.target = "_blank";		
		document.frm0904.submit();
	}

	function GenerateEnvelope(){
		//Print recipient
		var temp = document.frm0904.Recipient.value.split(":");
		document.frm0904.action = "../TPL/PrintEnvelope.asp?RecipientType=" + temp[0] + "&To=" + temp[1];		
		document.frm0904.target = "_blank";
		document.frm0904.submit();

		document.frm0904.CC.value = "";
		for (var i = 0; i < document.frm0904.CCList.options.length; i++) {
			if (document.frm0904.CCList.options[i].selected) {
				document.frm0904.CC.value = document.frm0904.CC.value + ":" + document.frm0904.CCList.options[i].value;
			}
		}
		
		if (document.frm0904.CC.value.length > 0) {
			document.frm0904.CC.value = document.frm0904.CC.value.substring(1, document.frm0904.CC.value.length);
		}
		
		//Print CCs
		for (var i = 0; i < document.frm0904.CCList.options.length; i++) {
			if (document.frm0904.CCList.options[i].selected) {
				document.frm0904.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0904.CCList.options[i].value;
				document.frm0904.target = "_blank";
				document.frm0904.submit();
			}
		}

		//if CC Client
		if (document.frm0904.CCClient.checked) {
			document.frm0904.action = "../TPL/PrintEnvelope.asp?RecipientType=Client&To=<%=Request.QueryString("intAdult_id")%>";
			document.frm0904.target = "_blank";
			document.frm0904.submit();
		}				
	}
	</script>
</head>
<body onLoad="document.frm0904.Subject.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0904">
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
		rsClient.Requery();
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
		while (!rsClient.EOF) {
		%>
			<option value="Client:<%=(rsClient.Fields.Item("intAdult_Id").Value)%>" <%=((rsLetter.Fields.Item("chvRx_Class").Value=="Client")?"SELECTED":"")%>><%=(rsClient.Fields.Item("chvName").Value)%></option>
		<%
			rsClient.MoveNext();
		}
		while (!rsContact.EOF) {
		%>
			<option value="Contact:<%=(rsContact.Fields.Item("intContact_id").Value)%>" <%=(((rsLetter.Fields.Item("chvRx_Class").Value=="Contact")&&(rsContact.Fields.Item("intContact_id").Value==rsLetter.Fields.Item("intRecipient_id").Value))?"SELECTED":"")%>><%=rsContact.Fields.Item("chvName").Value%> (<%=(rsContact.Fields.Item("chvRelationship").Value)%>)
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
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" value="<%=rsLetter.Fields.Item("chvLetter_Name").Value%>" maxlength="50" size="30" tabindex="6"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=FilterDate(rsLetter.Fields.Item("dtsSend_Date").Value)%>" size="11" maxlength="10" tabindex="7" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="8" accesskey="L">
			<option value="0" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="Canada_Post")?"SELECTED":"")%>>Canada Post
			<option value="1" <%=((rsLetter.Fields.Item("chvSend_Method").Value=="e-Mail")?"SELECTED":"")%>>E-Mail
		</select></td>
	</tr>	
</table>
<hr>
<textarea name="CustomLetterContent" cols="90" rows="50" tabindex="9" readonly>
<%=Trim(rsLetter.Fields.Item("chvNote").Value)%>
</textarea>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="View Letter" onClick="Save();" tabindex="10" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m001q0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>';" tabindex="11" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="CC">
</form>
</body>
</html>
<%
rsClient.Close();
rsContact.Close();
%>