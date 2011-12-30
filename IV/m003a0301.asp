<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_insert")) == "true") {
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");		
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_Insert_Notes(" + Request.QueryString("intEquip_Set_id") + ",'13','" + Notes+ "',"+ Session("insStaff_id") + ",0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	Response.Redirect("InsertSuccessful.html");
}
%>
<html>
<head>
	<title>New Notes</title>
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
		if (!CheckTextArea(document.frm0301.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
		document.frm0301.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0301.Notes.focus()" >
<form name="frm0301" method="POST" action="<%=MM_editAction%>">
<h5>New Notes</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="5" accesskey="F" tabindex="1"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
    	<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="3" onClick="window.close();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
</form>
</body>
</html>