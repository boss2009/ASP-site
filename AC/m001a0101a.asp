<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("State"))=="Search") {
	var rsClient__chvFilter = "";
	if(String(Request.QueryString("chvFilter")) != "undefined") {
		rsClient__chvFilter = String(Request.QueryString("chvFilter"));
	}
	var rsClient = Server.CreateObject("ADODB.Recordset");
	rsClient.ActiveConnection = MM_cnnASP02_STRING;
	rsClient.Source = "{call dbo.cp_Adult_Client2E(1,0,'"+ rsClient__chvFilter.replace(/'/g, "''") + "')}";
	rsClient.CursorType = 0;
	rsClient.CursorLocation = 2;
	rsClient.LockType = 3;
	rsClient.Open();
}
%>
<html>
<head>
	<title>Search For Existing Client</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="JavaScript" src="../js/m001Srh01.js"></script>
	<script language="JavaScript" src="../js/MyFunctions.js"></script>
	<script language="JavaScript">
	if (window.focus) self.focus();
	function openWindow(page, name){
		if (page!='nothing') win1=window.open(page, "", "width=750,height=500,scrollbars=1,left=20,top=20,status=1");
		return ;
	}

	function CnstrFltr(output) {
		var stgFilter = ACfltr_01(document.frm0101.StringSearchOperand.value,"",document.frm0101.StringSearchOperator.value,document.frm0101.StringSearchTextOne.value,"");
		document.frm0101.action = "m001a0101a.asp?chvFilter=" + stgFilter ;
		document.frm0101.State.value = "Search";
		document.frm0101.submit() ;
	}

	function ViewClient(){
		adult_id = document.frm0101.ClientsFound.value;
		if (adult_id > 0) {
			openWindow('m001FS3.asp?intAdult_id='+adult_id);
		} else {
			alert("Select a client.");
			document.frm0101.ClientsFound.focus();
			return ;
		}
	}

	function CreateNew() {
		document.frm0101.action="m001a0101b.asp";
		document.frm0101.submit();
	}

	function EditClient() {
		var client_id = document.frm0101.ClientsFound.value;
		if (client_id > 0) {
			document.frm0101.action="m001FS3.asp?intAdult_id=" + client_id;
			document.frm0101.submit();
		} else {
			alert("Select a client.");
			document.frm0101.ClientsFound.focus();
			return ;
		}
	}

	function Init(){
	<%
	if (String(Request.Form("State"))=="Search") {
	%>
		document.frm0101.New.disabled = false;
		document.frm0101.Edit.disabled = false;
		document.frm0101.ClientsFound.focus();
	<%
	} else {	
	%>
		document.frm0101.StringSearchOperand.focus();
	<%
	}
	%>
	}
	</script>
</head>
<body onload="Init();">
<form name="frm0101" method="POST" action="">
<h5>Search for Existing Client</h5>
<hr>
<table cellpadding="1" cellspacing="1">
    <tr>
		<td nowrap>
			<select name="StringSearchOperand" onChange="document.frm0101.StringSearchTextOne.value='';" tabindex="1" accesskey="F">
				<option value="11">Last Name
				<option value="12">First Name				
<!--			<option value="">Sin-->
			</select>
			<select name="StringSearchOperator" tabindex="2">
				<option value="1">starts with
				<option value="2">contains
				<option value="3">is												
				<option value="4">ends with				
			</select>
			<input type="text" name="StringSearchTextOne" tabindex="3">
			<input type="button" value="Search" onClick="CnstrFltr(1);" tabindex="5" class="btnstyle">
			<input type="button" value="Cancel" onClick="window.close();" tabindex="6" class="btnstyle">
		</td>
    </tr>
</table><br>
&nbsp;<span style="font-family: Courier;"><%=FormatContact("Last Name","First Name","SIN","")%></span>
<select name="ClientsFound" size="20" style="width: 400px; height: 280px; font-family: Courier;" ondblclick="ViewClient();" tabindex="7" accesskey="L">
<%
if (String(Request.Form("State"))=="Search") {
	while (!rsClient.EOF) {
%>
		<option value="<%=rsClient.Fields.Item("intAdult_id").Value%>"><%=FormatContact(rsClient.Fields.Item("chvLst_Name").Value,rsClient.Fields.Item("chvFst_Name").Value,FormatSIN(rsClient.Fields.Item("chrSIN_no").Value),"")%>
<%
		rsClient.MoveNext();
	}
}
%>
</select>
<br><br>
To create a new client, click <input type="button" name="New" value="New Client" disabled onClick="CreateNew();" tabindex=8" class="btnstyle"><br>
-OR-<br>
Highlight one of the above client, and click <input type="button" name="Edit" value="Edit" disabled class="btnstyle" onClick="EditClient();">
<input type="hidden" name="MM_flag" value="false">
<input type="hidden" name="MM_curOprd">
<input type="hidden" name="MM_curOptr">
<input type="hidden" name="State">
</form>
</body>
</html>