<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_insert")) == "true") {
	var rsReminderNote = Server.CreateObject("ADODB.Recordset");
	rsReminderNote.ActiveConnection = MM_cnnASP02_STRING;
	rsReminderNote.Source = "{call dbo.cp_pjt_RmdLst2("+Session("insStaff_id")+",0,"+Request.Form("Function")+","+Session("insStaff_id")+","+Request.Form("Recipient")+",'"+String(Request.Form("Subject")).replace(/'/g, "''")+"','"+String(Request.Form("Notes")).replace(/'/g, "''")+"','A',0,"+Session("MM_UserAuthorization")+",0)}";
	rsReminderNote.CursorType = 0;
	rsReminderNote.CursorLocation = 2;
	rsReminderNote.LockType = 3;
	rsReminderNote.Open();
	Response.Redirect("AddDeleteSuccessful.asp?action=Add");
}

var rsFunction = Server.CreateObject("ADODB.Recordset");
rsFunction.ActiveConnection = MM_cnnASP02_STRING;
rsFunction.Source = "{call dbo.cp_ASP_Lkup(702)}";
rsFunction.CursorType = 0;
rsFunction.CursorLocation = 2;
rsFunction.LockType = 3;
rsFunction.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_Lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
%>
<html>
<head>
	<title>New Reminder</title>
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
	function AlertTimeOut(){
		alert("Session will timeout in 5 minutes.  Please save your work.");		
	}
	
	function Save(){
		if (Trim(document.frm0101.Subject.value)=="") {
			if (!confirm("Save without Subject?")) return ;
		}		
		if (!CheckTextArea(document.frm0101.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
		if (!CheckDate(document.frm0101.CreatedOn.value)) {
			alert("Invalid Create Date.");
			document.frm0101.CreatedOn.focus();
			return ;
		}
		document.frm0101.submit();
	}
	</script>	
</head>
<body onLoad="document.frm0101.Recipient.focus();setTimeout('AlertTimeOut()',2100000);">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0101">
<h5>New Reminder</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td nowrap>Recipient:</td>
		<td nowrap><select name="Recipient" tabindex="1" accesskey="F">
			<% 
			while (!rsStaff.EOF) { 
			%>
				<option value="<%=(rsStaff.Fields.Item("insStaff_id").Value)%>"><%=(rsStaff.Fields.Item("chvname").Value)%></option>
			<% 
				rsStaff.MoveNext();
			}
			%>
        </select></td>
    </tr>
    <tr> 
		<td nowrap>Subject:</td>
		<td nowrap><input type="text" name="Subject" maxlength="50" size="40" tabindex="2"></td>
    </tr>
    <tr> 
		<td nowrap>Function:</td>
		<td nowrap><select name="Function" tabindex="3">
			<% 
			while (!rsFunction.EOF) { 
			%>
				<option value="<%=(rsFunction.Fields.Item("insFTNid").Value)%>"><%=(rsFunction.Fields.Item("intMODno").Value)%>.<%=(rsFunction.Fields.Item("insFSTno").Value)%>.<%=(rsFunction.Fields.Item("insFSTsubno").Value)%>:<%=(rsFunction.Fields.Item("ncvFTNname").Value)%> (<%=rsFunction.Fields.Item("ncvMODname").Value%>)</option>
			<%
				rsFunction.MoveNext();
			}
			%>
		</select></td>
	</tr>
    <tr> 
		<td nowrap>Created On:</td>
		<td nowrap>
			<input type="text" name="CreatedOn" maxlength="10" size="11" tabindex="4" readonly value="<%=CurrentDate()%>" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
    </tr>
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap><textarea name="Notes" cols="80" rows="20" tabindex="5" accesskey="L"></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="6" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="7" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>
<%
rsFunction.Close();
rsStaff.Close();
%>