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
	var rsEquipmentRepairStatus = Server.CreateObject("ADODB.Recordset");
	rsEquipmentRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
	rsEquipmentRepairStatus.Source = "{call dbo.cp_eqpsrv_repsts("+Request.Form("MM_recordId")+","+Request.Form("RepairStatus")+",0,'E',0)}";
	rsEquipmentRepairStatus.CursorType = 0;
	rsEquipmentRepairStatus.CursorLocation = 2;
	rsEquipmentRepairStatus.LockType = 3;
	rsEquipmentRepairStatus.Open();
	Response.Redirect("UpdateSuccessful.asp?page=m009e0104.asp&intEquip_srv_id="+Request.QueryString("intEquip_srv_id"))
}

var rsEquipmentRepairStatus = Server.CreateObject("ADODB.Recordset");
rsEquipmentRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentRepairStatus.Source = "{call dbo.cp_eqpsrv_repsts("+ Request.QueryString("intEquip_Srv_id") + ",0,1,'Q',0)}";
rsEquipmentRepairStatus.CursorType = 0;
rsEquipmentRepairStatus.CursorLocation = 2;
rsEquipmentRepairStatus.LockType = 3;
rsEquipmentRepairStatus.Open();

var rsRepairStatus = Server.CreateObject("ADODB.Recordset");
rsRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
rsRepairStatus.Source = "{call dbo.cp_repair_status(0,'',0,'Q',0)}";
rsRepairStatus.CursorType = 0;
rsRepairStatus.CursorLocation = 2;
rsRepairStatus.LockType = 3;
rsRepairStatus.Open();
%>
<html>
<head>
	<title>Update Repair Status</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
	<script language="javascript" src="../js/MyFunctions.js"></script>
	<script for="document" event="onkeyup()" language="JavaScript">
	if (window.event.ctrlKey) {
		switch (window.event.keyCode) {
			case 83 :
				//alert("S");
				Save();
			break;
			case 85:
				//alert("U");
				document.frm0105.reset();
			break;
			case 76 :
				//alert("L");
				window.location.href='m009e0101.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>';
			break;
		}
	}
	</script>		
	<script language="Javascript">
	function Save(){
		document.frm0105.submit();
	}
	</script>	
</head>
<body onLoad="javascript:document.frm0105.RepairStatus.focus()">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="frm0105">
<h5>Repair Status</h5>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td nowrap>Repair Status:</td>
		<td nowrap><select name="RepairStatus" tabindex="1" accesskey="F">
		<% 
		while (!rsRepairStatus.EOF) {
		%>
			<option value="<%=(rsRepairStatus.Fields.Item("insEq_Repair_Sts_id").Value)%>" <%=((rsRepairStatus.Fields.Item("insEq_Repair_Sts_id").Value == rsEquipmentRepairStatus.Fields.Item("insRepair_Status").Value)?"SELECTED":"")%> ><%=(rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value)%>
		<%
			rsRepairStatus.MoveNext();
		}
		%>
		</select></td>
    </tr>
</table>	
<hr>
<table cellpadding="1" cellspacing="1">
	<tr> 
		<td><input type="button" value="Save" tabindex="2" onClick="Save();" class="btnstyle"></td>
		<td><input type="reset" value="Undo Changes" tabindex="3" class="btnstyle"></td>
		<td><input type="button" value="Close" tabindex="4" onClick="window.location.href='m009e0101.asp?intEquip_Srv_id=<%=Request.QueryString("intEquip_Srv_id")%>'" class="btnstyle"></td>
	</tr>
</table>
<input type="hidden" name="MM_update" value="true">
<input type="hidden" name="MM_recordId" value="<%= rsEquipmentRepairStatus.Fields.Item("intEquip_Srv_id").Value %>">
</form>
</body>
</html>
<%
rsEquipmentRepairStatus.Close();
rsRepairStatus.Close();
%>