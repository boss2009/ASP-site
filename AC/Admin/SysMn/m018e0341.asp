<%@language="JAVASCRIPT"%>
<!--#include file="../../inc/ASPUtility.inc" -->
<!--#include file="../../Connections/cnnASP02.asp" -->
<!--#include file="../../inc/ASPCheckAdminLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var IsLoanDocument = ((Request.Form("IsLoanDocument")=="1") ? "1":"0");	
	var IsOutstandingDocument = ((Request.Form("IsOutstandingDocument")=="1") ? "1":"0");	
	var IsDeclineDocument = ((Request.Form("IsDeclineDocument")=="1") ? "1":"0");	
	var IsPendingDocument = ((Request.Form("IsPendingDocument")=="1") ? "1":"0");	
	var IncludeEquipment = ((Request.Form("IncludeEquipment")=="1") ? "1":"0");	
	var TemplateName = String(Request.Form("TemplateName")).replace(/'/g, "''");	
	var FileName = String(Request.Form("FileName")).replace(/'/g, "''");		
	var rsLetterTemplate = Server.CreateObject("ADODB.Recordset");
	rsLetterTemplate.ActiveConnection = MM_cnnASP02_STRING;
	rsLetterTemplate.Source = "{call dbo.cp_Letter_template("+ Request.Form("MM_recordId") + ","+Request.Form("TemplateType")+",'" + TemplateName + "',"+Request.Form("DocumentType")+",'" + FileName + "'," + IsLoanDocument + "," + IsOutstandingDocument + "," + IsDeclineDocument + "," + IsPendingDocument + "," + IncludeEquipment + "," + Session("insStaff_id")+",0,'E',0)}";
	rsLetterTemplate.CursorType = 0;
	rsLetterTemplate.CursorLocation = 2;
	rsLetterTemplate.LockType = 3;
	rsLetterTemplate.Open();
	Response.Redirect("m018q0341.asp");
}

var rsLetterTemplate = Server.CreateObject("ADODB.Recordset");
rsLetterTemplate.ActiveConnection = MM_cnnASP02_STRING;
rsLetterTemplate.Source = "{call dbo.cp_Letter_template("+Request.QueryString("insTemplate_id")+",0,'',0,'',0,0,0,0,0,0,1,'Q',0)}";
rsLetterTemplate.CursorType = 0;
rsLetterTemplate.CursorLocation = 2;
rsLetterTemplate.LockType = 3;
rsLetterTemplate.Open();
%>
<html>
<head>
	<title>Update Letter Template Lookup</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0341.reset();
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
		if (Trim(document.frm0341.TemplateName.value)==""){
			alert("Enter Template Name.");
			document.frm0341.TemplateName.focus();
			return ;		
		}
		document.frm0341.submit();
	}
	</script>
</head>
<body onLoad="document.frm0341.TemplateName.focus();">
<form name="frm0341" method="POST" action="<%=MM_editAction%>">
<h5>Update Template Lookup</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td>Template Name:</td>
		<td><input type="text" name="TemplateName" value="<%=(rsLetterTemplate.Fields.Item("chvTemplate_Name").Value)%>" tabindex="1" accesskey="F" ></td>
    </tr>
    <tr> 
		<td>Template Type:</td>
		<td><select name="TemplateType" tabindex="2">
				<option value="0" <%=((rsLetterTemplate.Fields.Item("bitIs_Form_Ltr").Value == 0)?"SELECTED":"")%>>Letter
				<option value="1" <%=((rsLetterTemplate.Fields.Item("bitIs_Form_Ltr").Value == 1)?"SELECTED":"")%>>Form
		</select></td>
    </tr>
    <tr> 
		<td>Document Type:</td>
		<td><select name="DocumentType" tabindex="3">		
				<option value="0" <%=((rsLetterTemplate.Fields.Item("insDocType").Value=="0")?"SELECTED":"")%>>Others		
				<option value="1" <%=((rsLetterTemplate.Fields.Item("insDocType").Value=="1")?"SELECTED":"")%>>Accept
				<option value="2" <%=((rsLetterTemplate.Fields.Item("insDocType").Value=="2")?"SELECTED":"")%>>PILAT
				<option value="3" <%=((rsLetterTemplate.Fields.Item("insDocType").Value=="3")?"SELECTED":"")%>>Decline
				<option value="4" <%=((rsLetterTemplate.Fields.Item("insDocType").Value=="4")?"SELECTED":"")%>>Pending
		</select></td>
    </tr>
    <tr> 
		<td>File Name:</td>
		<td><input type="textbox" name="FileName" value="<%=rsLetterTemplate.Fields.Item("chvFileName").Value%>" tabindex="4"></td> 
    </tr>
    <tr> 
		<td>Is Loan Document:</td>
		<td><input type="checkbox" name="IsLoanDocument" <%=((rsLetterTemplate.Fields.Item("bitIs_Loan_Doc").Value == 1)?"CHECKED":"")%> value="1" tabindex="5" class="chkstyle"></td>
	</tr>	
    <tr> 
		<td>Is Outstanding Document:</td>
		<td><input type="checkbox" name="IsOutstandingDocument" <%=((rsLetterTemplate.Fields.Item("bitIs_OutStand_Doc").Value == 1)?"CHECKED":"")%> value="1" tabindex="6" class="chkstyle"></td>
	</tr>	
    <tr> 
		<td>Is Decline Document:</td>
		<td><input type="checkbox" name="IsDeclineDocument" <%=((rsLetterTemplate.Fields.Item("bitIs_Decline_Doc").Value == 1)?"CHECKED":"")%> value="1" tabindex="7" class="chkstyle"></td>
	</tr>	
    <tr> 
		<td>Is Pending Document:</td>
		<td><input type="checkbox" name="IsPendingDocument" <%=((rsLetterTemplate.Fields.Item("bitIs_Pending_Doc").Value == 1)?"CHECKED":"")%> value="1" tabindex="8" class="chkstyle"></td>
	</tr>	
    <tr> 
		<td>Include Equipment:</td>
		<td><input type="checkbox" name="IncludeEquipment" <%=((rsLetterTemplate.Fields.Item("bitIs_Include_Eqp").Value == 1)?"CHECKED":"")%> value="1" tabindex="9" accesskey="L" class="chkstyle"></td>
	</tr>		
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="10" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="11" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="12" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsLetterTemplate.Fields.Item("insTemplate_id").Value%>">
</form>
</body>
</html>
<%
rsLetterTemplate.Close();
%>