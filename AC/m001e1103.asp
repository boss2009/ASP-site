<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var ActionRequired = String(Request.Form("ActionRequired")).replace(/'/g, "''");	
	var rsFollowUp = Server.CreateObject("ADODB.Recordset");
	rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
	rsFollowUp.Source="{call dbo.cp_follow_up("+Request.Form("MM_recordId")+",'3',"+ Request.QueryString("intAdult_id") +",0,'',0,0,0,'','',0,0,0,0,'',0,'','','',0,'',0,0,'',0,0,0,'','','"+Request.Form("FollowUpDate")+"','"+Request.Form("RequiredBy")+"','"+Request.Form("DateCompleted")+"','"+ActionRequired+"',0,'E',0)}";
	rsFollowUp.CursorType = 0;
	rsFollowUp.CursorLocation = 2;
	rsFollowUp.LockType = 3;
	rsFollowUp.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m001q1103.asp&intAdult_id="+Request.QueryString("intAdult_id"));
}

var rsFollowUp = Server.CreateObject("ADODB.Recordset");
rsFollowUp.ActiveConnection = MM_cnnASP02_STRING;
rsFollowUp.Source = "{call dbo.cp_Follow_up("+ Request.QueryString("intFlwup_id") + ",'3',0,0,'',0,0,0,'','',0,0,0,0,'',0, '','','',0,'',0.00,0.00,'',0,0,0,'','','','','','',1,'Q',0)}";
rsFollowUp.CursorType = 0;
rsFollowUp.CursorLocation = 2;
rsFollowUp.LockType = 3;
rsFollowUp.Open();
%>
<html>
<head>
	<title>Update General Follow-Up</title>
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
				document.frm1103.reset();
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
		if (!CheckDate(document.frm1103.FollowUpDate.value)) {
			alert("Invalid Follow-Up Date.");
			document.frm1103.FollowUpDate.focus();
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
<body onLoad="document.frm1103.FollowUpDate.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm1103">
<h5>Update General Follow-Up</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Date:</td>
		<td nowrap>
			<input type="text" name="FollowUpDate" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsGFup_date").Value)%>" size="11" maxlength="10" tabindex="1" accesskey="F" onChange="FormatDate(this)" >
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Action Required:</td>
		<td nowrap valign="top"><textarea name="ActionRequired" rows="5" cols="65" tabindex="2"><%=(rsFollowUp.Fields.Item("chvActn_Req").Value)%></textarea></td>
    </tr>
    <tr> 
		<td nowrap>Required By:</td>
		<td nowrap>
			<input type="text" name="RequiredBy" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsRqtBy").Value)%>" size="11" maxlength="10" tabindex="3" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap>Date Completed:</td>
		<td nowrap>
			<input type="text" name="DateCompleted" value="<%=FilterDate(rsFollowUp.Fields.Item("dtsCmplt_at").Value)%>" size="11" maxlength="10" tabindex="4" accesskey="L" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" name="Reset" value="Undo Changes" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="7" onClick="history.back()" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%=rsFollowUp.Fields.Item("intFlwup_id").Value%>">
</form>
</body>
</html>
<%
rsFollowUp.Close();
%>