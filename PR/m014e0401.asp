<!--------------------------------------------------------------------------
* File Name: m014e0401.asp
* Title: Edit Notes
* Main SP: cp_purchase_requisition_note
* Description: This page updates requested/received notes.
* Author: T.H
--------------------------------------------------------------------------->
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
	var rsNotes__Notes_type = ((String(Request.Form("NotesType")) == "Requested")?"21":"22");
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");			
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_Purchase_Requisition_Note(0,'"+Notes+"',"+Request.QueryString("int_Note_id")+","+ Session("insStaff_id")+",'"+rsNotes__Notes_type+"',0,'E',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	rsNotes.Close();
	Response.Redirect("UpdateSuccessful.asp?page=m014q0401.asp&insPurchase_Req_id="+Request.QueryString("insPurchase_Req_id"));
}

var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_Purchase_Requisition_Note("+Request.QueryString("insPurchase_Req_id")+",'',"+Request.QueryString("int_Note_id")+",0,'',1,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();
%>
<html>
<head>
	<title><%=((Request.QueryString("NotesType")=="Requested")?"Requested":"Received")%> Notes</title>
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
				document.frm0401.reset();
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
		if (!CheckTextArea(document.frm0401.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
		document.frm0401.submit();
	}
	</script>
</head>
<body onLoad="document.frm0401.Notes.focus();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0401">
<h5><%=((Request.QueryString("NotesType")=="Requested")?"Requested":"Received")%> Notes</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td valign="top">Notes:</td>
		<td valign="top"><textarea name="Notes" cols="65" rows="10" tabindex="1" accesskey="F"><%=rsNotes.Fields.Item("chvNote_Desc").Value%></textarea></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>	  
		<td><input type="button" value="Close" tabindex="4" onClick="history.back();" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_Update" value="true">
</form>
</body>
</html>
<%
rsNotes.Close();
%>