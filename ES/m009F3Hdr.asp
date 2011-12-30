<%@language="JAVASCRIPT"%>
<!--#include file="../inc/ASPUtility.inc" -->
<!--#include file="../Connections/cnnASP02.asp" -->
<!--#include file="../inc/ASPCheckLogin.inc" -->
<%
var rsEquipmentServiceHeader = Server.CreateObject("ADODB.Recordset");
rsEquipmentServiceHeader.ActiveConnection = MM_cnnASP02_STRING;

// + Nov.04.2005
//rsEquipmentServiceHeader.Source = "{call dbo.cp_FrmHdr_9A("+Request.QueryString("intEquip_Srv_id")+",0)}";
rsEquipmentServiceHeader.Source = "{call dbo.cp_FrmHdr_9("+Request.QueryString("intEquip_Srv_id")+",0)}";

rsEquipmentServiceHeader.CursorType = 0;
rsEquipmentServiceHeader.CursorLocation = 2;
rsEquipmentServiceHeader.LockType = 3;
rsEquipmentServiceHeader.Open();

var rsEquipmentRepairStatus = Server.CreateObject("ADODB.Recordset");
rsEquipmentRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
rsEquipmentRepairStatus.Source = "{call dbo.cp_eqpsrv_repsts("+ Request.QueryString("intEquip_Srv_id") + ",0,1,'Q',0)}";
rsEquipmentRepairStatus.CursorType = 0;
rsEquipmentRepairStatus.CursorLocation = 2;
rsEquipmentRepairStatus.LockType = 3;
rsEquipmentRepairStatus.Open();

var rsRepairStatus = Server.CreateObject("ADODB.Recordset");
rsRepairStatus.ActiveConnection = MM_cnnASP02_STRING;
rsRepairStatus.Source = "{call dbo.cp_repair_status("+rsEquipmentRepairStatus.Fields.Item("insRepair_Status").Value+",'',1,'Q',0)}";
rsRepairStatus.CursorType = 0;
rsRepairStatus.CursorLocation = 2;
rsRepairStatus.LockType = 3;
rsRepairStatus.Open();
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>					
function DoTrim(str, side)							
	dim strRet								
	strRet = str								
										
	If (side = 0) Then						
		strRet = LTrim(str)						
	ElseIf (side = 1) Then						
		strRet = RTrim(str)						
	Else									
		strRet = Trim(str)						
	End If									
										
	DoTrim = strRet								
End Function									
</SCRIPT>									
<html>
<head>
	<title>Equipment Service Header</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="stylesheet" href="../css/MyStyle.css" type="text/css">
</head>
<body>
<div class="TestPanel" style="width: 570px;"> 
<%
if (!rsEquipmentServiceHeader.EOF) {
	switch (rsEquipmentServiceHeader.Fields.Item("insEq_user_type").Value) {
		//client
		case 3:
%>
<table cellspacing="1" cellpadding="1">
	<tr> 
		<td valign="top" nowrap><b>Inventory Name:</b></td>
		<td valign="top" width="210"><%=(rsEquipmentServiceHeader.Fields.Item("chvInventory_Name").Value)%></td>
		<td valign="top" nowrap><b>Inventory ID:</b></td>
		<td valign="top" nowrap><%=ZeroPadFormat(rsEquipmentServiceHeader.Fields.Item("intEquip_Set_id").Value,8)%></td>
    </tr>
    <tr> 
		<td nowrap valign="top"><b>User Name:</b></td>
		<td nowrap valign="top"><%=(rsEquipmentServiceHeader.Fields.Item("chvUsr_Name").Value)%></td>
		<td nowrap valign="top"><b>Date Requested:</b></td>
		<td nowrap valign="top"><%=FilterDate(rsEquipmentServiceHeader.Fields.Item("dtsRequested_date").Value)%></td>
    </tr>
	<tr>
		<td nowrap valign="top"><b>Case Manager:</b></td>
		<td nowrap valign="top"><%=(rsEquipmentServiceHeader.Fields.Item("chvCaseManager").Value)%></td>
		<td nowrap valign="top"><b>Repair Status:</b></td>
		<td nowrap valign="top"><%=rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value%></td>	
	</tr>	
<!--
    <tr> 
		<td nowrap valign="top"><b>Funding Source:</b></td>
		<td nowrap valign="top" colspan="3"><%=Trim(rsEquipmentServiceHeader.Fields.Item("chvfunding_source_name").Value)%></td>
    </tr>
-->
</table>
<%
		break;
		//school
		case 4:
%>
<table cellspacing="1" cellpadding="1" border="0">
    <tr> 
		<td valign="top" nowrap><b>Inventory Name:</b></td>
		<td valign="top" width="210"><%=(rsEquipmentServiceHeader.Fields.Item("chvInventory_Name").Value)%></td>
		<td valign="top" nowrap><b>Inventory ID:</b></td>
		<td valign="top" nowrap><%=ZeroPadFormat(rsEquipmentServiceHeader.Fields.Item("intEquip_Set_id").Value,8)%></td>
    </tr>
    <tr> 
		<td valign="top" nowrap><b>User Name:</b></td>
		<td valign="top" nowrap><%=(rsEquipmentServiceHeader.Fields.Item("chvSch_Name").Value)%></td>
		<td valign="top" nowrap><b>Date Requested:</b></td>
		<td valign="top" nowrap><%=FilterDate(rsEquipmentServiceHeader.Fields.Item("dtsRequested_date").Value)%></td>
    </tr>
	<tr>
		<td valign="top" nowrap><b>Case Manager:</b></td>
		<td valign="top" nowrap><%=(rsEquipmentServiceHeader.Fields.Item("chvCaseManager").Value)%></td>
		<td valign="top" nowrap><b>Repair Status:</b></td>
		<td valign="top" nowrap><%=rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value%></td>	
	</tr>
<!--	
    <tr> 
		<td valign="top" nowrap><b>Funding Source:</b></td>
		<td valign="top"><%=(rsEquipmentServiceHeader.Fields.Item("chvfunding_source_name").Value)%></td>
    </tr>
-->
</table>
<%
		break;
		//staff
		case 1:
%>
<table cellspacing="1" cellpadding="1">
	<tr> 
		<td valign="top" nowrap><b>Inventory Name:</b></td>
		<td valign="top" width="210"><%=(rsEquipmentServiceHeader.Fields.Item("chvInventory_Name").Value)%></td>
		<td valign="top" nowrap><b>Inventory ID:</b></td>
		<td valign="top" nowrap><%=ZeroPadFormat(rsEquipmentServiceHeader.Fields.Item("intEquip_Set_id").Value,8)%></td>
    </tr>
    <tr> 
		<td nowrap valign="top"><b>User Name:</b></td>
		<td nowrap valign="top"><%=(rsEquipmentServiceHeader.Fields.Item("chvUsr_Name").Value)%></td>
		<td nowrap valign="top"><b>Date Requested:</b></td>
		<td nowrap valign="top"><%=FilterDate(rsEquipmentServiceHeader.Fields.Item("dtsRequested_date").Value)%></td>
    </tr>
    <tr> 
		<td></td>
		<td></td>
		<td nowrap valign="top"><b>Repair Status:</b></td>
		<td nowrap valign="top"><%=rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value%></td>
    </tr>
</table>
<%		
		break;
		case 0:
%>
<table cellspacing="1" cellpadding="1">
	<tr> 
		<td valign="top" nowrap><b>Inventory Name:</b></td>
		<td valign="top" width="210"><%=(rsEquipmentServiceHeader.Fields.Item("chvInventory_Name").Value)%></td>
		<td valign="top" nowrap><b>Inventory ID:</b></td>
		<td valign="top" nowrap><%=ZeroPadFormat(rsEquipmentServiceHeader.Fields.Item("intEquip_Set_id").Value,8)%></td>
    </tr>
    <tr> 
		<td nowrap valign="top"><b>User Name:</b></td>
		<td nowrap valign="top">No User</td>
		<td nowrap valign="top"><b>Date Requested:</b></td>
		<td nowrap valign="top"><%=FilterDate(rsEquipmentServiceHeader.Fields.Item("dtsRequested_date").Value)%></td>
    </tr>
    <tr> 
		<td></td>
		<td></td>
		<td nowrap valign="top"><b>Repair Status:</b></td>
		<td nowrap valign="top"><%=rsRepairStatus.Fields.Item("chvEq_Repair_Sts_Desc").Value%></td>
    </tr>
</table>
<%		
		break;
		default:
%>
<i>Information not available for this Equipment Service.</i><br>
<br>
<br>
<br>
<br>
<br>
<%		
		break;
	}
} else {
%>
<i>Information not available for this Equipment Service.</i><br>
<br>
<br>
<br>
<br>
<br>
<%
}
%>
</div>
</body>
</html>
<%
rsEquipmentServiceHeader.Close();
%>