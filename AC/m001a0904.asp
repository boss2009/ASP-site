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

	CC[9] = ((Request.Form("CCClient")=="on")?"99999":"0");

	var DocumentName = String(Request.Form("DocumentName")).replace(/'/g, "''");	
	var CustomLetterContent = String(Request.Form("CustomLetterContent")).replace(/'/g, "''");	
	var rsTemplate = Server.CreateObject("ADODB.Recordset");
	rsTemplate.ActiveConnection = MM_cnnASP02_STRING;
	rsTemplate.Source = "{call dbo.cp_insert_crspltr_custom(0,0,"+Request.QueryString("intAdult_id")+",0,"+Session("insStaff_id")+",0,4,"+temp2[1]+","+CC[0]+","+CC[1]+","+CC[2]+","+CC[3]+","+CC[4]+","+CC[5]+","+CC[6]+","+CC[7]+","+CC[8]+","+CC[9]+",'"+DocumentName+"',"+Is_Recipient_Client+",'"+Request.Form("DateGenerated")+"',"+Request.Form("MailMethod")+",'"+CustomLetterContent+"',0)}";
	rsTemplate.CursorType = 0;
	rsTemplate.CursorLocation = 2;
	rsTemplate.LockType = 3;
	rsTemplate.Open();
	Response.Redirect("InsertSuccessful.html");
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

		document.frm0904.CC.value = "";
		
		for (var i = 0; i < document.frm0904.CCList.options.length; i++) {
			if (document.frm0904.CCList.options[i].selected) document.frm0904.CC.value = document.frm0904.CC.value + ":" + document.frm0904.CCList.options[i].value;
		}
		if (document.frm0904.CC.value.length > 0) document.frm0904.CC.value = document.frm0904.CC.value.substring(1, document.frm0904.CC.value.length);
		
		var temp = document.frm0904.action;
				
		if (document.frm0904.MailMethod.value=="0") {		
			if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();			
			document.frm0904.action = "../TPL/CustomLetterTemplate.asp";
		} else {
			document.frm0904.action = "../TPL/E-CustomLetterTemplate.asp";
		}		

		document.frm0904.target = "_blank";		
		document.frm0904.submit();
		
		document.frm0904.action = temp;
		document.frm0904.target = "_self";
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
		
		//if CC client	
		if (document.frm0904.CCClient.checked) {
			document.frm0904.action = "../TPL/PrintEnvelope.asp?RecipientType=Client&To=<%=Request.QueryString("intAdult_id")%>";
			document.frm0904.target = "_blank";
			document.frm0904.submit();
		}			
	}
		
	function ChangeType(){
		if (document.frm0904.Type.value == "4") {	
	<%
	if (String(Request.QueryString("intBuyout_req_id")) != "undefined") {
	%>
			window.location.href = "m001a0902.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
	<%
	} else if (String(Request.QueryString("intLoan_req_id")) != "undefined") {
	%>
			window.location.href = "m001a0903.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";
	<%
	} else {	
	%>
			window.location.href = "m001a0901.asp?intAdult_id=<%=Request.QueryString("intAdult_id")%>";
	<%
	}
	%>
		}
	}
	</script>
</head>
<body onLoad="document.frm0904.Subject.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0904">
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
		rsClient.Requery();
		%>
		</select></td>
    </tr>
	<tr>
		<td nowrap>Type:</td>
		<td nowrap><select name="Type" tabindex="2" onChange="ChangeType();">
			<option value="4">Form Letter
			<option value="0" SELECTED>Custom Letter
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
		<td nowrap>Document Name:</td>
		<td nowrap><input type="text" name="DocumentName" maxlength="50" size="30" tabindex="6"></td>
    </tr>
    <tr> 
		<td nowrap>Date Generated:</td>
		<td nowrap>
			<input type="text" name="DateGenerated" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="7" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
	<tr>
		<td nowrap>Method:</td>
		<td nowrap><select name="MailMethod" tabindex="8" accesskey="L">
			<option value="0">Canada Post
			<option value="1">E-Mail
		</select></td>
	</tr>	
</table>
<hr>
<textarea name="CustomLetterContent" cols="90" rows="50" tabindex="9"></textarea>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Generate Letter" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" class="btnstyle"></td>
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
%>