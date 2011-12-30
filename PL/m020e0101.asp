<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request("MM_update")) == "true") {
	var rsReminderList = Server.CreateObject("ADODB.Recordset");
	rsReminderList.ActiveConnection = MM_cnnASP02_STRING;
	rsReminderList.Source = "{call dbo.cp_pjt_RmdLst2("+Session("insStaff_id")+","+Request.QueryString("intRmdLst_id")+",0,"+Session("insStaff_id")+",0,'','"+String(Request.Form("Notes")).replace(/'/g, "''")+"','E',0,"+Session("MM_UserAuthorization")+",0)}";
	rsReminderList.CursorType = 0;
	rsReminderList.CursorLocation = 2;
	rsReminderList.LockType = 3;
	rsReminderList.Open();
	Response.Redirect("m020q0101.asp");
}

var rsReminderList = Server.CreateObject("ADODB.Recordset");
rsReminderList.ActiveConnection = MM_cnnASP02_STRING;
rsReminderList.Source = "{call dbo.cp_pjt_RmdLst2(0,"+Request.QueryString("intRmdLst_id")+",0,0,0,'','','Q',1,"+Session("MM_UserAuthorization")+",0)}";
rsReminderList.CursorType = 0;
rsReminderList.CursorLocation = 2;
rsReminderList.LockType = 3;
rsReminderList.Open();
%>
<html>
<head>
	<title>Update Reminder List</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83:
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0101.reset();
			break;
			case 76 :
				//alert("L");
				window.location.href='m020q0101.asp';
			break;
		}
	}
	</script>	
	<script language="Javascript">	
	function AlertTimeOut(){
		alert("Session will timeout in 5 minutes.  Please save your work.");		
	}
	
	function Save(){
		if (!CheckTextArea(document.frm0101.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		document.frm0101.submit();
	}
	
	function Init(){
		document.frm0101.Subject.focus();
		setTimeout('AlertTimeOut()',2100000);
	}
	</script>
</head>
<body onLoad="Init();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0101">
<h5>Update Reminder Note</h5>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Subject:</td>
		<td nowrap valign="top"><input type="text" name="Subject" value="<%=(rsReminderList.Fields.Item("chvSubject").Value)%>" readonly tabindex="1" size="40" accesskey="F"></td>
	</tr>
	<tr>
		<td nowrap valign="top">Function:</td>
		<td nowrap valign="top"><input type="text" name="Function" value="<%=rsReminderList.Fields.Item("chvMod_name").Value%> in <%=rsReminderList.Fields.Item("chvGroup").Value%>" readonly tabindex="2" size="40"></td>
	</tr>
	<tr>
		<td nowrap valign="top">Created On:</td>
		<td nowrap valign="top"><input type="text" name="CreatedOn" value="<%=FilterDate(rsReminderList.Fields.Item("dtsRec_Create_date").Value)%>" readonly tabindex="3" size="20"></td>
	</tr>
	<tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="80" rows="20" tabindex="4" accesskey="L"><%=(rsReminderList.Fields.Item("chvNote").Value)%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="6" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.location.href='m020q0101.asp';" tabindex="7" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
</form>
</body>
</html>
<%
rsReminderList.Close();
%>