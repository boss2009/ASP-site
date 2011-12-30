<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("Transfer")) == "true") {
	var rsTransferClass = Server.CreateObject("ADODB.Recordset");
	rsTransferClass.ActiveConnection = MM_cnnASP02_STRING;
	rsTransferClass.Source = "{call dbo.cp_Transfer_Eqp_Class("+Request.Form("SubAbstractClass")+","+Request.Form("ToAbstractClass")+","+Session("insStaff_id")+",0)}";
	rsTransferClass.CursorType = 0;
	rsTransferClass.CursorLocation = 2;
	rsTransferClass.LockType = 3;
	rsTransferClass.Open();
	Response.Redirect("AddDeleteSuccessful2.asp?action=Transfer");
}

var rsAbstractClass = Server.CreateObject("ADODB.Recordset");
rsAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW(0,'A',0)}";
rsAbstractClass.CursorType = 0;
rsAbstractClass.CursorLocation = 2;
rsAbstractClass.LockType = 3;
rsAbstractClass.Open();

var rsSubAbstractClass__ClassID = ((String(Request.Form("AbstractClass")) != "undefined")?Request.Form("AbstractClass"):rsAbstractClass.Fields.Item("insEquip_Class_id").Value);

var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW("+rsSubAbstractClass__ClassID+",'S',0)}";
rsSubAbstractClass.CursorType = 0;
rsSubAbstractClass.CursorLocation = 2;
rsSubAbstractClass.LockType = 3;
rsSubAbstractClass.Open();
%>
<html>
<head>
	<title>Sub Abstract Class Transfer</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function SelectAbstract(){
		document.frm01t.submit();
	}
	
	function TransferClass(){
		document.frm01t.Transfer.value="true";
		document.frm01t.submit();
	}
	</script>	
</head>
<body>
<form name="frm01t" method="POST" action="m007t01.asp">
<h3>Sub Abstract Class Transfer</h3>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Transfer:</td>
		<td><select name="AbstractClass" accesskey="F" tabindex="1" onChange="SelectAbstract();" style="width: 200px">
			<% 
			while (!rsAbstractClass.EOF){ 
			%>
				<option value=<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%> <%=((Request.Form("AbstractClass")==rsAbstractClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsAbstractClass.Fields.Item("chvName").Value%>
			<%
				rsAbstractClass.MoveNext();
			}
			rsAbstractClass.MoveFirst();
			%>		
		</select></td>
		<td><select name="SubAbstractClass" tabindex="2" style="width: 200px">	
			<%
			while (!rsSubAbstractClass.EOF){ 
			%>
				<option value=<%=rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value%>><%=rsSubAbstractClass.Fields.Item("chvName").Value%>
			<%
				rsSubAbstractClass.MoveNext();
			}
			%>		
		</select></td>		
	</tr>
	<tr height="10">
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td>Under This Class:</td>	
		<td><select name="ToAbstractClass" accesskey="L" tabindex="3" style="width: 200px">
			<% 
			while (!rsAbstractClass.EOF){ 
			%>
				<option value=<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%>><%=rsAbstractClass.Fields.Item("chvName").Value%>
			<%
				rsAbstractClass.MoveNext();
			}
			%>				
		</select></td>
		<td>&nbsp;</td>
	</tr>	
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Transfer" onClick="TransferClass();" tabindex="4" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="Transfer" value="false">
</form>
</body>
</html>
<%
rsAbstractClass.Close();
rsSubAbstractClass.Close();
%>