<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../inc/ASPCheckLogin.inc"-->
<!--#include file="../Connections/cnnASP02.asp" -->
<%
if (String(Request.Form("Transfer")) == "true") {
	var rsTransferClass = Server.CreateObject("ADODB.Recordset");
	rsTransferClass.ActiveConnection = MM_cnnASP02_STRING;
	rsTransferClass.Source = "{call dbo.cp_Transfer_Eqp_Class("+Request.Form("ConcreteClass")+","+Request.Form("ToSubAbstractClass")+","+Session("insStaff_id")+",0)}";
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

var rsSubAbstractClass__ClassID = ((String(Request.Form("Initialized")) == "true")?Request.Form("AbstractClass"):rsAbstractClass.Fields.Item("insEquip_Class_id").Value);

var rsSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW("+rsSubAbstractClass__ClassID+",'S',0)}";
rsSubAbstractClass.CursorType = 0;
rsSubAbstractClass.CursorLocation = 2;
rsSubAbstractClass.LockType = 3;
rsSubAbstractClass.Open();

var rsToSubAbstractClass__ClassID = ((String(Request.Form("Initialized")) == "true")?Request.Form("ToAbstractClass"):rsAbstractClass.Fields.Item("insEquip_Class_id").Value);

var rsToSubAbstractClass = Server.CreateObject("ADODB.Recordset");
rsToSubAbstractClass.ActiveConnection = MM_cnnASP02_STRING;
rsToSubAbstractClass.Source = "{call dbo.cp_Eqp_Class_LW("+rsToSubAbstractClass__ClassID+",'S',0)}";
rsToSubAbstractClass.CursorType = 0;
rsToSubAbstractClass.CursorLocation = 2;
rsToSubAbstractClass.LockType = 3;
rsToSubAbstractClass.Open();

if ((String(Request.Form("Initialized")) == "true") && (String(Request.Form("CInitialize")) == "false")) {
	var rsConcreteClass = Server.CreateObject("ADODB.Recordset");
	rsConcreteClass.ActiveConnection = MM_cnnASP02_STRING;
	rsConcreteClass.Source = "{call dbo.cp_Eqp_Class_LW("+Request.Form("SubAbstractClass")+",'C',0)}";
	rsConcreteClass.CursorType = 0;
	rsConcreteClass.CursorLocation = 2;
	rsConcreteClass.LockType = 3;
	rsConcreteClass.Open();	
}
%>
<html>
<head>
	<title>Concrete Class Transfer</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="Javascript">
	function SelectClass(){
		document.frm02t.submit();
	}
	
	function TransferClass(){
		document.frm02t.Transfer.value="true";
		document.frm02t.submit();
	}
	</script>	
</head>
<body>
<form name="frm02t" method="POST" action="m007t02.asp">
<h3>Concrete Class Transfer</h3>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td>Transfer:</td>
		<td><select name="AbstractClass" accesskey="F" tabindex="1" onChange="document.frm02t.CInitialize.value='true';SelectClass();" style="width: 200px">
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
		<td><select name="SubAbstractClass" tabindex="2" style="width: 200px" onChange="SelectClass();">	
			<%
			while (!rsSubAbstractClass.EOF){ 
			%>
				<option value=<%=rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value%> <%=((Request.Form("SubAbstractClass")==rsSubAbstractClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsSubAbstractClass.Fields.Item("chvName").Value%>
			<%
				rsSubAbstractClass.MoveNext();
			}
			rsSubAbstractClass.MoveFirst();
			%>		
		</select></td>		
		<td><select name="ConcreteClass" tabindex="3" style="width: 200px">	
		<%
		if ((String(Request.Form("Initialized")) == "true") && (String(Request.Form("CInitialize")) == "false")) {
			while (!rsConcreteClass.EOF){ 
		%>
				<option value=<%=rsConcreteClass.Fields.Item("insEquip_Class_id").Value%> <%=((Request.Form("ConcreteClass")==rsConcreteClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsConcreteClass.Fields.Item("chvName").Value%>
		<%
				rsConcreteClass.MoveNext();
			}
		} else {
		%>
				<option value="">Select Sub Abstract Class
		<%
		}		
		%>		
		</select></td>				
	</tr>
	<tr height="10">
		<td colspan="4"></td>
	</tr>
	<tr>
		<td nowrap>Under This Class:</td>
		<td><select name="ToAbstractClass" tabindex="4" style="width: 200px" onChange="SelectClass();">
			<% 
			while (!rsAbstractClass.EOF) { 
			%>
				<option value=<%=rsAbstractClass.Fields.Item("insEquip_Class_id").Value%> <%=((Request.Form("ToAbstractClass")==rsAbstractClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsAbstractClass.Fields.Item("chvName").Value%>
			<%
				rsAbstractClass.MoveNext();
			}
			%>
		</select></td>
		<td><select name="ToSubAbstractClass" tabindex="5" style="width: 200px" accesskey="L">
			<%
			while (!rsToSubAbstractClass.EOF) {
			%>
				<option value=<%=rsToSubAbstractClass.Fields.Item("insEquip_Class_id").Value%> <%=((Request.Form("ToSubAbstractClass")==rsToSubAbstractClass.Fields.Item("insEquip_Class_id").Value)?" SELECTED":"")%>><%=rsToSubAbstractClass.Fields.Item("chvName").Value%>
			<%
				rsToSubAbstractClass.MoveNext();
			}
			%>
		</select></td>
		<td></td>
	</tr>
</table>
<hr>
<table cellpadding="1" cellspacing="1">
	<tr>
		<td><input type="button" value="Transfer" onClick="TransferClass();" tabindex="6" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="Transfer" value="false">
<input type="hidden" name="Initialized" value="true">
<input type="hidden" name="CInitialize" value="false">
</form>
</body>
</html>
<%
rsAbstractClass.Close();
rsSubAbstractClass.Close();
rsToSubAbstractClass.Close();
if ((String(Request.Form("Initialized")) == "true") && (String(Request.Form("CInitialize")) == "false")) rsConcreteClass.Close();
%>