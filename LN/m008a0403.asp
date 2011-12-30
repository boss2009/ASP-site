<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_Insert"))=="true") {
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_eqp_loaned_notes("+Request.QueryString("intLoan_req_id")+",0,'"+Notes+"',"+ Session("insStaff_id")+",'2',0,'A',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	rsNotes.Close();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New Equipment Loaned Notes</title>
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
				self.close();
			break;
		}
	}
	</script>
	<script language="Javascript">
	function Save(){
		if (!CheckTextArea(document.frm0403.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		document.frm0403.submit();
	}
	</script>
</head>
<body onLoad="document.frm0403.Notes.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0403">
<h5>New Equipment Loaned Notes</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="10" tabindex="1" accesskey="F"></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="3" onClick="self.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Insert" value="true">
</form>
</body>
</html>