<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#INCLUDE File="../inc/ASPCheckLogin.inc" -->
<!--#INCLUDE File="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
	MM_editAction += "?" + Request.QueryString;
}

var rsInventory = Server.CreateObject("ADODB.Recordset");
rsInventory.ActiveConnection = MM_cnnASP02_STRING;
rsInventory.Source = "select intIvtry_note_id from tbl_equip_inventory where intEquip_Set_id = " + Request.QueryString("intEquip_Set_id");
rsInventory.CursorType = 0;
rsInventory.CursorLocation = 2;
rsInventory.LockType = 3;
rsInventory.Open();

var MM_action = "";

if (rsInventory.Fields.Item("intIvtry_note_id").Value == null) {
	MM_action = "insert";
} else {
	MM_action = "update";
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_Get_Ivtry_Notes(0," + rsInventory.Fields.Item("intIvtry_note_id").Value + ",1,0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();	
}

if (String(Request.Form("MM_action")) == "update"){
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");		
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_Update_Notes(" + rsInventory.Fields.Item("intIvtry_note_id").Value + ",'" + Notes + "'," + Session("insStaff_id") + ",0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	Response.Redirect("UpdateSuccessful2.asp?page=m003e0301.asp&intEquip_Set_id="+Request.QueryString("intEquip_Set_id"));
}

if (String(Request.Form("MM_action")) == "insert"){
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");		
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_Insert_Notes(" + Request.QueryString("intEquip_Set_id") + ",'13','" + Notes+ "',"+ Session("insStaff_id") + ",0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();	
	Response.Redirect("UpdateSuccessful2.asp?page=m003e0301.asp&intEquip_Set_id="+Request.QueryString("intEquip_Set_id"));
}
%>
<html>
<head>
	<title>Notes</title>
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
<body onLoad="document.frm0301.Notes.focus();">
<form action="<%=MM_editAction%>" method="POST" name="frm0301">
<h5>Inventory Notes</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><textarea name="Notes" rows="15" cols="80" tabindex="1" accesskey="F"><%=((MM_action=="update")?rsNotes.Fields.Item("chvNote_Desc").Value:"")%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" onClick="Save();" tabindex="2" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>		
    </tr>
</table>
<input type="hidden" name="MM_action" value="<%=MM_action%>">
<input type="hidden" name="intEquip_Set_id" value="<%=Request.QueryString("intEquip_Set_id")%>">
</form>
</body>
</html>