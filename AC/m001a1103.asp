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
	var ActionRequired = String(Request.Form("ActionRequired")).replace(/'/g, "''");	
	var rsFollowUp = Server.CreateObject("ADODB.Recordset");
	rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
	rsFollowUp.Source="{call dbo.cp_follow_up(0,'3',"+ Request.QueryString("intAdult_id") +",0,'',0,0,0,'','',0,0,0,0,'',0,'','','',0,'',0,0,'',0,0,0,'','','"+Request.Form("GeneralFollowUpDate")+"','"+Request.Form("RequiredBy")+"','"+Request.Form("DateCompleted")+"','"+ActionRequired+"',0,'A',0)}";
	rsFollowUp.CursorType = 0;
	rsFollowUp.CursorLocation = 2;
	rsFollowUp.LockType = 3;
	rsFollowUp.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New General Follow-Up</title>
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
		if (!CheckDate(document.frm1103.GeneralFollowUpDate.value)) {
			alert("Invalid Follow-Up Date.");
			document.frm1103.GeneralFollowUpDate.focus();
			return ;
		}
		if (!CheckDate(document.frm1103.RequiredBy.value)){
			alert("Invalid Required By Date.");
			document.frm1103.RequiredBy.focus();
			return ;
		}
		if (!CheckDate(document.frm1103.DateCompleted.value)){
			alert("Invalid Completed Date.");
			document.frm1103.DateCompleted.focus();
			return ;
		}
		document.frm1103.submit();		
	}	
	</script>
</head>
<body onLoad="document.frm1103.GeneralFollowUpDate.focus();">
<form name="frm1103" method="POST" action="<%=MM_editAction%>">
<h5>New General Follow-Up</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Date:</td>
		<td nowrap>
			<input type="text" name="GeneralFollowUpDate" value="<%=CurrentDate()%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Action Required:</td>
		<td nowrap valign="top"><textarea name="ActionRequired" rows="5" cols="65" tabindex="2"></textarea></td>
    </tr>
    <tr> 
		<td nowrap>Required by:</td>
		<td nowrap>
			<input type="text" name="RequiredBy" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
    </tr>
    <tr> 
		<td nowrap>Date Completed:</td>
		<td nowrap>
			<input type="text" name="DateCompleted" size="11" maxlength="10" tabindex="4" accesskey="L" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>			
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="6" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>