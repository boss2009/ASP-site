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
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_pilat_srv_note("+Request.QueryString("insSchool_id")+",0,'"+Request.Form("Date")+"',0,0,"+Session("insStaff_id")+",'"+Notes+"','"+Request.Form("NoteTypeHexCode")+"',1,'A',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	Response.Redirect("InsertSuccessful.html");
}	

var rsNotesType = Server.CreateObject("ADODB.Recordset");
rsNotesType.ActiveConnection = MM_cnnASP02_STRING;
//rsNotesType.Source = "{call dbo.cp_service_type(0,0,0,2)}";
rsNotesType.Source = "select * from tbl_service_type where bitis_req_service = 0 and bitis_School_Class = 1 and bitis_active = 1 order by chvname asc";
rsNotesType.CursorType = 0;
rsNotesType.CursorLocation = 2;
rsNotesType.LockType = 3;
rsNotesType.Open();

var rsStaff = Server.CreateObject("ADODB.Recordset");
rsStaff.ActiveConnection = MM_cnnASP02_STRING;
rsStaff.Source = "{call dbo.cp_ASP_lkup(10)}";
rsStaff.CursorType = 0;
rsStaff.CursorLocation = 2;
rsStaff.LockType = 3;
rsStaff.Open();
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
		var temp = "";
		for (var i=0; i < document.frm0502.NoteType.length; i++){
			if (document.frm0502.NoteType[i].selected) {
				temp = temp + PadDecToHex(document.frm0502.NoteType[i].value);
			}
		}
		var zero = 40 - temp.length;
		for (var j = 0; j < zero; j++){
			temp = temp + String("0");
		}
		document.frm0502.NoteTypeHexCode.value=temp;
		document.frm0502.submit();
	}
	</script>
</head>
<body onLoad="javascript:document.frm0502.Date.focus()" >
<form action="<%=MM_editAction%>" method="POST" name="frm0502">
<h5>New Notes</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Date:</td>
		<td nowrap>
			<input type="textbox" name="Date" value="<%=CurrentDate()%>" accesskey="F" tabindex="1" size="11" maxlength="10" onChange="FormatDate(this)">
			<span style="font-size: 7pt">(mm/dd/yyyy)</span>
		</td>
	</tr>
    <tr>
		<td nowrap valign="top">Note Type:</td>
		<td nowrap><select name="NoteType" MULTIPLE size="8" tabindex="2">
			<%
			while (!rsNotesType.EOF) {
			%>
				<option value="<%=rsNotesType.Fields.Item("insService_type_id").Value%>"><%=rsNotesType.Fields.Item("chvname").Value%>
			<%
				rsNotesType.MoveNext();
			}
			%>		
		</select></td>
    </tr>
    <tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" rows="8" cols="65" tabindex="3" accesskey="L"></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" onClick="Save();" tabindex="4" class="btnstyle"></td>
		<td><input type="button" value="Close" onClick="window.close();" tabindex="5" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_insert" value="true">
<input type="hidden" name="NoteTypeHexCode">
</form>
</body>
</html>
<%
rsNotesType.Close();
rsStaff.Close();
%>