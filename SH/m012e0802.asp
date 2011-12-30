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

var rsDocumentCondition = Server.CreateObject("ADODB.Recordset");
rsDocumentCondition.ActiveConnection = MM_cnnASP02_STRING;
rsDocumentCondition.Source = "{call dbo.cp_Doc_Cdn_Rsn2(0,19,'',2,'Q',0)}";
rsDocumentCondition.CursorType = 0;
rsDocumentCondition.CursorLocation = 2;
rsDocumentCondition.LockType = 3;
rsDocumentCondition.Open();
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
	if (rsTemplate.Fields.Item("chvFileName").Value == "m012tpl002.asp") {
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
		document.frm0802.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0802.Recipient.value;
		document.frm0802.target = "_blank";
		document.frm0802.submit();

		document.frm0802.CC.value = "";
		for (var i = 0; i < document.frm0802.CCList.options.length; i++) {
			if (document.frm0802.CCList.options[i].selected) document.frm0802.CC.value = document.frm0802.CC.value + ":" + document.frm0802.CCList.options[i].value;
		}
		
		if (document.frm0802.CC.value.length > 0) document.frm0802.CC.value = document.frm0802.CC.value.substring(1, document.frm0802.CC.value.length);
		
		//Print CCs
		for (var i = 0; i < document.frm0802.CCList.options.length; i++) {
			if (document.frm0802.CCList.options[i].selected) {
				document.frm0802.action = "../TPL/PrintEnvelope.asp?RecipientType=Contact&To=" + document.frm0802.CCList.options[i].value;
				document.frm0802.target = "_blank";
				document.frm0802.submit();
			}
		}		
	}
		
	function Save(){
		if (!CheckDate(document.frm0802.DateGenerated.value)){
			alert("Invalid Date Generated.");
			document.frm0802.DateGenerated.focus();
			return ;
		}
		if (Trim(document.frm0802.DocumentName.value)=="") {
			alert("Enter Document Name.");
			document.frm0802.DocumentName.focus();
			return ;
		}
	
		if (confirm("Do you wish to generate envelopes?")) GenerateEnvelope();
			
		document.frm0802.CC.value = "";
		
		for (var i = 0; i < document.frm0802.CCList.options.length; i++) {
			if (document.frm0802.CCList.options[i].selected) document.frm0802.CC.value = document.frm0802.CC.value + ":" + document.frm0802.CCList.options[i].value;
		}

		if (document.frm0802.CC.value.length > 0) document.frm0802.CC.value = document.frm0802.CC.value.substring(1, document.frm0802.CC.value.length);
		
		if (document.frm0802.TransactionType.value=="Buyout") {
			if (document.frm0802.MailMethod.value=="0") {		
				document.frm0802.action = "../TPL/"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
			} else {
				document.frm0802.action = "../TPL/E-"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";			
			}
		} else {
			if (document.frm0802.MailMethod.value=="0") {
				document.frm0802.action = "../TPL/"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";		
			} else {
				document.frm0802.action = "../TPL/E-"+DocumentArray[document.frm0802.Template.selectedIndex][2]+"?insSchool_id=<%=Request.QueryString("insSchool_id")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>";			
			}
		}
		
		document.frm0802.target = "_blank";
		document.frm0802.submit();
	}
	
	function ChangeType(){
		if (document.frm0802.Type.value == "0") {
			window.location.href = "m012a0803.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>&Type=<%=Request.QueryString("Type")%>&intLoan_req_id=<%=Request.QueryString("intLoan_req_id")%>&intBuyout_req_id=<%=Request.QueryString("intBuyout_req_id")%>";
		}
	}
	</script>
</head>
<body onLoad="document.frm0802.Type.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0802">
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
		if (rsTemplate.Fields.Item("chvFileName").Value == "m012tpl002.asp") {		
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
<h5>PILAT Accept</h5>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><b>Type of Referral:</b></td>
		<td><select name="PilatAcceptReferralType" tabindex="8">
			<%
			if (Request.QueryString("TransactionType")=="Loan") {
			%>						
				<option value="1" <%=((rsLetter.Fields.Item("insINT_01").Value==1)?"SELECTED":"")%>>Low Utilization
				<option value="2" <%=((rsLetter.Fields.Item("insINT_01").Value==2)?"SELECTED":"")%>>Interim
				<option value="3" <%=((rsLetter.Fields.Item("insINT_01").Value==3)?"SELECTED":"")%>>Donation
			<%
			} else {
			%>				
				<option value="4" <%=((rsLetter.Fields.Item("insINT_01").Value==4)?"SELECTED":"")%>>Purchase
			<%
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td><b>Equipment List:</b></td>
		<td><select name="EquipmentList" tabindex="9">
				<option value="0" <%=((rsLetter.Fields.Item("insINT_02").Value==0)?"SELECTED":"")%>>None
			<%
			if (Request.QueryString("TransactionType")=="Loan") {
			%>								
				<option value="1" <%=((rsLetter.Fields.Item("insINT_02").Value==1)?"SELECTED":"")%>>Loan Equipment			
				<option value="2" <%=((rsLetter.Fields.Item("insINT_02").Value==2)?"SELECTED":"")%>>Donation Equipment
			<%
			} else {
			%>								
				<option value="3" <%=((rsLetter.Fields.Item("insINT_02").Value==3)?"SELECTED":"")%>>Buyout Equipment
			<%
			}
			%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherEquipmentList" value="<%=rsLetter.Fields.Item("chvText_01").Value%>" tabindex="10" maxlength="80" size="60"></td>
	</tr>
	<tr>
		<td><b>Equipment Conditions:</b></td>
		<td></td>
	</tr>
<%
if (Request.QueryString("TransactionType") == "Loan") {
%>	
	<tr>
		<td></td>
		<td>
			<input type="radio" name="EquipmentConditions" value="1" class="chkstyle" <%=((rsLetter.Fields.Item("insINT_03").Value==1)?"CHECKED":"")%> tabindex="11">Low Utilization - Loan Review Date&nbsp;<input type="text" name="LoanReviewDate" value="<%=FilterDate(rsLetter.Fields.Item("dtsDate_01").Value)%>" maxlength="10" size="12" onChange="FormatDate(this)"><span style="font-size: 7pt">&nbsp;(mm/dd/yyyy)</span>
		</td>	
	</tr>	
	<tr>
		<td></td>
		<td>
			<input type="radio" name="EquipmentConditions" value="2" class="chkstyle" <%=((rsLetter.Fields.Item("insINT_03").Value==2)?"CHECKED":"")%> tabindex="12">Interim Loan - Return Date&nbsp;<input type="text" name="ReturnDate" value="<%=FilterDate(rsLetter.Fields.Item("dtsDate_02").Value)%>" maxlength="10" size="12" onChange="FormatDate(this)"><span style="font-size: 7pt">&nbsp;(mm/dd/yyyy)</span>
		</td>
	</tr>
	<tr>
		<td></td>
		<td><input type="radio" name="EquipmentConditions" class="chkstyle" value="3" <%=((rsLetter.Fields.Item("insINT_03").Value==3)?"CHECKED":"")%> tabindex="13">Donation</td>
	</tr>	
<%
} else {
%>
	<tr>
		<td></td>
		<td><input type="radio" name="EquipmentConditions" class="chkstyle" value="4" <%=((rsLetter.Fields.Item("insINT_03").Value==4)?"CHECKED":"")%> tabindex="14">Purchase</td>
	</tr>	
	<tr>
		<td align="right">Other</td>
		<td><input type="text" name="OtherEquipmentConditions" value="<%=rsLetter.Fields.Item("chvText_02").Value%>" maxlength="80" size="60" tabindex="16"></td>
	</tr>
<%
}
%>
	<tr>
		<td><b>Document Conditions:</b></td>
		<td><select name="DocumentConditionOne" tabindex="17">
				<option value="0">(none)
	<%	
	rsCondition.Requery();	
	if (!rsCondition.EOF) {
		while (!rsDocumentCondition.EOF) {
			if (rsDocumentCondition.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {
	%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
	<%	
			}		
			rsDocumentCondition.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionTwo" tabindex="18">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsDocumentCondition.Requery();	
		while (!rsDocumentCondition.EOF) {
			if (rsDocumentCondition.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
	<%	
			}		
			rsDocumentCondition.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionThree" tabindex="19">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsDocumentCondition.Requery();	
		while (!rsDocumentCondition.EOF) {
			if (rsDocumentCondition.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
	<%	
			}		
			rsDocumentCondition.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td></td>
		<td><select name="DocumentConditionFour" tabindex="20">
				<option value="0">(none)
	<%
	if (!rsCondition.EOF) {
		rsDocumentCondition.Requery();	
		while (!rsDocumentCondition.EOF) {
			if (rsDocumentCondition.Fields.Item("intDoc_id").Value==rsCondition.Fields.Item("intDoc_Id").Value) {	
	%>
				<option value="<%=rsDocumentCondition.Fields.Item("intDoc_id").Value%>" SELECTED><%=rsDocumentCondition.Fields.Item("chvDocDesc").Value%>
	<%	
			}		
			rsDocumentCondition.MoveNext();
		}
		rsCondition.MoveNext();
	}
	%>
		</select></td>
	</tr>
	<tr>
		<td align="right">Other:</td>
		<td><input type="text" name="OtherDocumentCondition" value="<%=Trim(rsLetter.Fields.Item("chvText_03").Value)%>" maxlength="80" size="60" tabindex="21"></td>
	</tr>
	<tr>
		<td colspan="2"><input type="checkbox" name="TrainingRequested" <%=((rsLetter.Fields.Item("bitCkBx_01").Value=="1")?"CHECKED":"")%> class="chkstyle" tabindex="22"><b>Training Requested</b></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="View Letter" onClick="Save();" tabindex="23" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m012q0801.asp?insSchool_id=<%=Request.QueryString("insSchool_id")%>';" tabindex="24" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="CC" value="">
<input type="hidden" name="TransactionType" value="<%=Request.QueryString("TransactionType")%>">
<input type="hidden" name="intLoan_req_id" value="<%=Request.QueryString("intLoan_req_id")%>">
<input type="hidden" name="intBuyout_req_id" value="<%=Request.QueryString("intBuyout_req_id")%>">
</form>
</body>
</html>
<%
rsInstitution.Close();
rsContact.Close();
rsTemplate.Close();
rsLetter.Close();
rsCC.Close();
rsCondition.Close();
%>