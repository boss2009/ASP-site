<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
var MM_editAction = Request.ServerVariables("URL");
if (Request.QueryString) {
  MM_editAction += "?" + Request.QueryString;
}

if (String(Request.Form("MM_Update"))=="true") {
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_buyout_sold_notes(0,"+Request.QueryString("int_Note_id")+",'"+Notes+"',"+ Session("insStaff_id")+",'6',0,'E',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	rsNotes.Close();
	Response.Redirect("UpdateSuccessful.asp?page=m010q0403.asp&intBuyout_req_id="+Request.QueryString("intBuyout_req_id"));
}

var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_buyout_sold_notes(0,"+Request.QueryString("int_Note_id")+",'',0,'',1,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();
%>
<html>
<head>
	<title>Update Equipment Sold Notes</title>
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
		if (!CheckTextArea(document.frm0402.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		document.frm0402.submit();
	}
	</script>
</head>
<body onLoad="document.frm0402.Notes.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0402">
<h5>Equipment Sold Notes</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" cols="65" rows="10" tabindex="1" accesskey="F"><%=rsNotes.Fields.Item("chvNote_Desc").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="3" onClick="history.back();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Update" value="true">
</form>
</body>
</html>
<%
rsNotes.Close();
%>