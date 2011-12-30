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
	var Notes = String(Request.Form("Notes")).replace(/'/g, "''");
	var rsNotes = Server.CreateObject("ADODB.Recordset");
	rsNotes.ActiveConnection = MM_cnnASP02_STRING;
	rsNotes.Source = "{call dbo.cp_pilat_srv_note("+Request.QueryString("insSchool_id")+","+Request.QueryString("intSrv_Note_id")+",'"+Request.Form("Date")+"',0,0,"+Request.Form("CreatedBy")+",'"+Notes+"','"+Request.Form("NoteTypeHexCode")+"',0,'E',0)}";
	rsNotes.CursorType = 0;
	rsNotes.CursorLocation = 2;
	rsNotes.LockType = 3;
	rsNotes.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m012q0501.asp&insSchool_id="+Request.QueryString("insSchool_id"));
}	

var rsNotes = Server.CreateObject("ADODB.Recordset");
rsNotes.ActiveConnection = MM_cnnASP02_STRING;
rsNotes.Source = "{call dbo.cp_pilat_srv_note("+Request.QueryString("insSchool_id")+","+Request.QueryString("intSrv_Note_id")+",'',0,0,0,'','',1,'Q',0)}";
rsNotes.CursorType = 0;
rsNotes.CursorLocation = 2;
rsNotes.LockType = 3;
rsNotes.Open();

var rsNoteType = Server.CreateObject("ADODB.Recordset");
rsNoteType.ActiveConnection = MM_cnnASP02_STRING;
//rsNotesType.Source = "{call dbo.cp_service_type(0,0,0,2)}";
rsNoteType.Source = "select * from tbl_service_type where bitis_req_service = 0 and bitis_School_Class = 1 and bitis_active = 1 order by chvname asc";
rsNoteType.CursorType = 0;
rsNoteType.CursorLocation = 2;
rsNoteType.LockType = 3;
rsNoteType.Open();

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
	<title>Update Notes</title>
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
				document.frm0502.reset();
			break;			
		   	case 76 :
				//alert("L");
				history.go("-1");
			break;
		}
	}	
	</script>	
	<script language="Javascript">	
	function Save(){
		if (!CheckTextArea(document.frm0502.Notes, 4000)){
			alert("Text area cannot exceed 4000 characters.");
			return ;
		}
	
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
	
	function HighLight(id){
		for (var i=0; i < document.frm0502.NoteType.length; i++){
			if (String(document.frm0502.NoteType[i].value)==String(id)) 
			document.frm0502.NoteType[i].selected = true;
		}
	}
	
	function Init(){
	<%
	while (!rsNotes.EOF) {
	%>
		HighLight('<%=rsNotes.Fields.Item("insNote_type_id").Value%>');		
	<%
		rsNotes.MoveNext();
	}
	rsNotes.MoveFirst();
	%>
		document.frm0502.Date.focus();
	}
	</script>
</head>
<body onLoad="Init();">
<form action="<%=MM_editAction%>" method="POST" name="frm0502">
<h5>Notes</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr> 
		<td nowrap>Date:</td>
		<td nowrap><input type="textbox" name="Date" value="<%=rsNotes.Fields.Item("dtsService_Date").Value%>" accesskey="F" tabindex="1" size="11" maxlength="10" onChange="FormatDate(this)"></td>
	</tr>
    <tr> 
		<td nowrap>Created By:</td>
		<td nowrap><select name="CreatedBy" tabindex="2">
			<%
			while (!rsStaff.EOF) {
			%>
				<option value="<%=rsStaff.Fields.Item("insStaff_id").Value%>" <%=((rsNotes.Fields.Item("insSrv_Staff_id").Value==rsStaff.Fields.Item("insStaff_id").Value)?"SELECTED":"")%>><%=rsStaff.Fields.Item("chvName").Value%>
			<%
				rsStaff.MoveNext();
			}
			%>				
		</select></td>
	</tr>
    <tr>
		<td nowrap valign="top">Note Type:</td>
		<td nowrap valign="top"><select name="NoteType" MULTIPLE size="8" tabindex="3">
			<%
			while (!rsNoteType.EOF) {
			%>
				<option value="<%=rsNoteType.Fields.Item("insService_type_id").Value%>"><%=rsNoteType.Fields.Item("chvname").Value%>
			<%
				rsNoteType.MoveNext();
			}
			%>		
		</select></td>
    </tr>
    <tr>
		<td nowrap valign="top">Notes:</td>
		<td nowrap valign="top"><textarea name="Notes" rows="8" cols="65" tabindex="4" accesskey="L"><%=rsNotes.Fields.Item("chvNotes").Value%></textarea></td>
    </tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Save" tabindex="5" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="6" class="btnstyle"></td>	  
		<td><input type="button" value="Close" tabindex="7" onClick="history.go(-1);" class="btnstyle"></td>
    </tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="NoteTypeHexCode">
</form>
</body>
</html>
<%
rsNotes.Close();
rsNoteType.Close();
rsStaff.Close();
%>